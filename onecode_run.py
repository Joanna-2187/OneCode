"""Run code blocks on the current OneNote page (Windows desktop) and write the output back.

A code block is a line starting with `#!<language>` plus the lines under it, up to the
first blank line (or a previous output block). Output goes in a monospace block right
under the code and is replaced on every run.

    python onecode_run.py            # run once on the current page
    python onecode_run.py --hotkey   # stay resident, run on Ctrl+Alt+R
    python onecode_run.py --dry-run  # run blocks, print results, don't touch the page

Needs: pip install pywin32, OneNote desktop open on the page, and `java` (11+) on PATH.
"""
import argparse
import ctypes
import html
import os
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from ctypes import wintypes

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
ET.register_namespace("one", NS)
Q = lambda tag: f"{{{NS}}}{tag}"

OUTPUT_META = "onecode.output"
MARKER = re.compile(r"^#!\s*(\w+)\s*$")
TIMEOUT = 10  # seconds per block
MAX_OUTPUT = 20_000  # characters written back


# ---------- reading the page ----------

def oe_text(oe):
    """Plain text of one OE: HTML from its <one:T> children, <br> as newlines."""
    parts = []
    for t in oe.findall(Q("T")):
        s = re.sub(r"<br\s*/?>", "\n", t.text or "", flags=re.I)
        s = re.sub(r"<[^>]+>", "", s)
        parts.append(html.unescape(s))
    s = "".join(parts).replace("\xa0", " ")
    # OneNote autocorrect turns quotes into curly ones; code needs straight ones.
    return s.translate({0x2018: "'", 0x2019: "'", 0x201C: '"', 0x201D: '"'})


def is_output(oe):
    return any(m.get("name") == OUTPUT_META for m in oe.findall(Q("Meta")))


def find_blocks(root):
    """Yield (parent OEChildren, [OEs in block], language, code, existing output OE or None)."""
    for children in root.iter(Q("OEChildren")):
        oes = list(children.findall(Q("OE")))
        i = 0
        while i < len(oes):
            lines = oe_text(oes[i]).split("\n")
            m = MARKER.match(lines[0].strip()) if not is_output(oes[i]) else None
            if not m:
                i += 1
                continue
            block, code_lines = [oes[i]], lines[1:]
            j = i + 1
            while j < len(oes) and not is_output(oes[j]) and oe_text(oes[j]).strip():
                block.append(oes[j])
                code_lines += oe_text(oes[j]).split("\n")
                j += 1
            out = oes[j] if j < len(oes) and is_output(oes[j]) else None
            yield children, block, m.group(1).lower(), "\n".join(code_lines), out
            i = j + (1 if out is not None else 0)


# ---------- running code ----------

def run_java(code, workdir):
    # Hoist imports; wrap bare statements in a class so snippets can be a few lines.
    if not re.search(r"\b(class|interface|record|enum)\s+\w+", code):
        imports = [l for l in code.splitlines() if l.strip().startswith("import ")]
        body = [l for l in code.splitlines() if not l.strip().startswith("import ")]
        code = "\n".join(imports) + "\npublic class Main {\n public static void main(String[] args) throws Exception {\n" \
               + "\n".join(body) + "\n }\n}\n"
    path = os.path.join(workdir, "Main.java")
    with open(path, "w", encoding="utf-8") as f:
        f.write(code)
    return ["java", path]  # source-launcher mode: compiles and runs, no javac step


def run_python(code, workdir):
    path = os.path.join(workdir, "snippet.py")
    with open(path, "w", encoding="utf-8") as f:
        f.write(code)
    return [sys.executable, "-I", path]


RUNNERS = {"java": run_java, "python": run_python, "py": run_python}


def run_block(lang, code):
    if lang not in RUNNERS:
        return f"No runner for '{lang}'. Available: {', '.join(sorted(RUNNERS))}"
    with tempfile.TemporaryDirectory() as workdir:
        try:
            cmd = RUNNERS[lang](code, workdir)
            p = subprocess.run(cmd, cwd=workdir, capture_output=True, text=True,
                               encoding="utf-8", errors="replace", timeout=TIMEOUT)
        except FileNotFoundError as e:
            return f"Can't start '{cmd[0]}': is it installed and on PATH? ({e})"
        except subprocess.TimeoutExpired:
            return f"Timed out after {TIMEOUT}s"
    out = p.stdout + (("\n" if p.stdout else "") + p.stderr if p.stderr else "")
    if p.returncode:
        out += f"\n[exit code {p.returncode}]"
    out = out.rstrip() or "(no output)"
    return out[:MAX_OUTPUT] + ("\n[truncated]" if len(out) > MAX_OUTPUT else "")


# ---------- writing output back ----------

def output_html(text):
    body = html.escape(text).replace(" ", "&nbsp;").replace("\r", "").replace("\n", "<br/>")
    return f"<span style='font-family:Consolas;font-size:10pt;color:#444444'>{body}</span>"


def set_output(children, last_code_oe, existing, text):
    if existing is None:
        existing = ET.Element(Q("OE"))  # no objectID: OneNote creates it
        ET.SubElement(existing, Q("Meta"), name=OUTPUT_META, content="1")
        ET.SubElement(existing, Q("T"))
        idx = list(children).index(last_code_oe) + 1
        children.insert(idx, existing)
    existing.find(Q("T")).text = output_html(text)


def parent_outline(root, element):
    for outline in root.iter(Q("Outline")):
        if any(e is element for e in outline.iter()):
            return outline


# ---------- OneNote COM ----------

def onenote_app():
    import win32com.client
    return win32com.client.Dispatch("OneNote.Application")


def run_page(dry_run=False):
    app = onenote_app()
    page_id = app.Windows.CurrentWindow.CurrentPageId
    root = ET.fromstring(app.GetPageContent(page_id, "", 0, 2))  # piBasic, xs2013
    touched, results = [], 0
    for children, block, lang, code, existing in list(find_blocks(root)):
        text = run_block(lang, code)
        results += 1
        print(f"--- {lang} ---\n{text}")
        if not dry_run:
            set_output(children, block[-1], existing, text)
            outline = parent_outline(root, children)
            if outline is not None and outline not in touched:
                touched.append(outline)
    if not results:
        print("No code blocks found (first line must be e.g. #!java).")
    elif not dry_run:
        page = ET.Element(Q("Page"), ID=page_id)
        page.extend(touched)
        app.UpdatePageContent(ET.tostring(page, encoding="unicode"))


# ---------- global hotkey ----------

def hotkey_loop():
    MOD_CONTROL, MOD_ALT, MOD_NOREPEAT, WM_HOTKEY = 0x2, 0x1, 0x4000, 0x312
    user32 = ctypes.windll.user32
    if not user32.RegisterHotKey(None, 1, MOD_CONTROL | MOD_ALT | MOD_NOREPEAT, ord("R")):
        sys.exit("Couldn't register Ctrl+Alt+R (already taken?).")
    print("Ctrl+Alt+R runs the current page. Ctrl+C to quit.")
    msg = wintypes.MSG()
    try:
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == WM_HOTKEY:
                try:
                    run_page()
                except Exception as e:  # keep the tool alive across OneNote hiccups
                    print(f"Run failed: {e}")
    finally:
        user32.UnregisterHotKey(None, 1)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--hotkey", action="store_true", help="stay resident, run on Ctrl+Alt+R")
    ap.add_argument("--dry-run", action="store_true", help="print results, don't edit the page")
    args = ap.parse_args()
    hotkey_loop() if args.hotkey else run_page(args.dry_run)
