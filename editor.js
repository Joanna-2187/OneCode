const LANGS = {
  javascript: "JavaScript", python: "Python", java: "Java",
  csharp: "C#", cpp: "C++", c: "C", rust: "Rust", r: "R", html: "HTML", css: "CSS",
  sql: "SQL", json: "JSON", plaintext: "Plain text",
};
const DEFAULT_LANG = "java";
// highlight.js name where it differs from Monaco's
const HLJS_NAME = { html: "xml", plaintext: "plaintext" };

// OneNote strips CSS classes, so token colours must be inlined.
const TOKEN_COLORS = {
  "hljs-keyword": "#0000ff", "hljs-built_in": "#267f99", "hljs-type": "#267f99",
  "hljs-literal": "#0000ff", "hljs-number": "#098658", "hljs-string": "#a31515",
  "hljs-regexp": "#811f3f", "hljs-comment": "#008000", "hljs-title": "#795e26",
  "hljs-function": "#795e26", "hljs-attr": "#e50000", "hljs-name": "#800000",
  "hljs-tag": "#800000", "hljs-variable": "#001080", "hljs-params": "#001080",
  "hljs-meta": "#af00db", "hljs-selector-class": "#800000", "hljs-property": "#001080",
};

// Starter text per language, in that language's own comment syntax (JSON has no comments).
const HINT = "Write code here, then click Insert into page";
const placeholder = (lang) => {
  if (lang === "json") return "";
  if (lang === "plaintext") return HINT + "\n";
  if (lang === "python" || lang === "r") return `# ${HINT}\n`;
  if (lang === "sql") return `-- ${HINT}\n`;
  if (lang === "html") return `<!-- ${HINT} -->\n`;
  if (lang === "css") return `/* ${HINT} */\n`;
  return `// ${HINT}\n`;
};

let editor;
const $ =(id) => document.getElementById(id);

function setStatus(msg) {
  $("status").textContent = msg;
  setTimeout(() => { if ($("status").textContent === msg) $("status").textContent = ""; }, 3000);
}

function inlineStyles(root) {
  root.querySelectorAll("span[class]").forEach((el) => {
    const cls = [...el.classList].find((c) => TOKEN_COLORS[c]);
    if (cls) el.style.color = TOKEN_COLORS[cls];
    if (el.classList.contains("hljs-comment")) el.style.fontStyle = "italic";
    el.removeAttribute("class");
  });
}

// Last successful run, so Insert can put the print output inside the block.
let lastRun = null; // { code, lang, text }

function buildHtml(code, lang, output) {
  const holder = document.createElement("div");
  const hl = HLJS_NAME[lang] || lang;
  const result = window.hljs.getLanguage(hl)
    ? window.hljs.highlight(code, { language: hl, ignoreIllegals: true }).value
    : window.hljs.highlight(code, { language: "plaintext" }).value;
  holder.innerHTML = result;
  inlineStyles(holder);
  // One table cell with a background is the most reliable "code block" in OneNote.
  // mso-no-proof + lang="zxx" ask Office to skip spell check (no squiggles) inside the block.
  const pre = "margin:0;font-family:Consolas,'Courier New',monospace;font-size:10pt;mso-no-proof:yes";
  // OneNote collapses whitespace even inside <pre>, so use divs with <br> and &nbsp;.
  const outHtml = output
    ? `<div lang="zxx" style="${pre};margin-top:6px;padding-top:6px;border-top:1px solid #ddd;color:#444"><b>Output:</b><br>${keepWhitespace(escapeHtml(output))}</div>`
    : "";
  return `<table style="border-collapse:collapse"><tr><td style="background:#f5f5f5;border:1px solid #ddd;padding:6px 10px">` +
    `<div lang="zxx" style="${pre};color:#1f1f1f">${keepWhitespace(holder.innerHTML)}</div>${outHtml}` +
    `</td></tr></table>`;
}

// Newlines -> <br>, spaces -> &nbsp;, in text only (never inside tags).
function keepWhitespace(html) {
  return html.split(/(<[^>]+>)/)
    .map((p) => (p.startsWith("<") ? p : p.replace(/\r?\n/g, "<br>").replace(/ /g, "&nbsp;")))
    .join("");
}

function escapeHtml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

async function insertIntoPage() {
  const code = editor.getValue();
  if (!code.trim()) return setStatus("Nothing to insert.");
  const lang = $("lang").value;
  // Runnable languages: run first if there's no up-to-date output, so prints land in the block.
  if (RUNNERS[lang] && !(lastRun && lastRun.code === code && lastRun.lang === lang)) await runCode();
  const ran = lastRun && lastRun.code === code && lastRun.lang === lang;
  const payload = { code, lang, output: ran ? lastRun.text : "" };
  // The pop-out dialog has no page access, so the pane does the actual insert.
  if (IS_DIALOG) Office.context.ui.messageParent(JSON.stringify({ type: "insert", ...payload }));
  else if (IN_OFFICE) doInsert(payload);
  else copyFormatted(payload, "Copied — paste into OneNote.");
}

// Same styled block that Insert puts on the page (colours, grey box, Output), as HTML on
// the clipboard, so it pastes formatted into OneNote desktop/web, Word, Outlook, etc.
// Plain text rides along for code editors and other plain-text targets.
async function copyFormatted({ code, lang, output }, msg = "Copied.") {
  const html = buildHtml(code, lang, output);
  const text = output ? `${code}\nOutput:\n${output}` : code;
  try {
    await navigator.clipboard.write([new ClipboardItem({
      "text/html": new Blob([html], { type: "text/html" }),
      "text/plain": new Blob([text], { type: "text/plain" }),
    })]);
    setStatus(msg);
  } catch { setStatus("Copy blocked."); }
}

function doInsert({ code, lang, output }) {
  const html = buildHtml(code, lang, output);
  Office.context.document.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, (r) => {
    setStatus(r.status === Office.AsyncResultStatus.Succeeded ? "Inserted." : "Failed: " + r.error.message);
  });
}

// ---- Pop-out window: same page opened in an Office dialog with ?dialog=1 ----
const IS_DIALOG = new URLSearchParams(location.search).has("dialog");
let dialog = null;

const state = () => ({ code: editor.getValue(), lang: $("lang").value, theme: $("theme").value });
function applyState(s) {
  $("lang").value = s.lang; $("theme").value = s.theme;
  monaco.editor.setModelLanguage(editor.getModel(), s.lang);
  monaco.editor.setTheme(s.theme);
  if (editor.getValue() !== s.code) editor.setValue(s.code);
}

function popOut() {
  if (dialog) return setStatus("Large window is already open.");
  const url = location.href.split("?")[0] + "?dialog=1";
  Office.context.ui.displayDialogAsync(url, { height: 45, width: 45, displayInIframe: false }, (r) => {
    if (r.status !== Office.AsyncResultStatus.Succeeded) return setStatus("Can't open window: " + r.error.message);
    dialog = r.value;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (a) => {
      const m = JSON.parse(a.message);
      if (m.type === "ready") dialog.messageChild(JSON.stringify(state()));
      else if (m.type === "sync") applyState(m);
      else if (m.type === "insert") { applyState(m); doInsert(m); }
      else if (m.type === "close") { applyState(m); dialog.close(); dialog = null; }
    });
    dialog.addEventHandler(Office.EventType.DialogEventReceived, () => { dialog = null; }); // closed with X
  });
}

function initDialog() {
  document.body.classList.add("dialog");
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (a) => applyState(JSON.parse(a.message)));
  const send = (type) => Office.context.ui.messageParent(JSON.stringify({ type, ...state() }));
  let t;
  editor.onDidChangeModelContent(() => { clearTimeout(t); t = setTimeout(() => send("sync"), 400); });
  $("lang").addEventListener("change", () => send("sync"));
  $("theme").addEventListener("change", () => send("sync"));
  $("done").onclick = () => send("close");
  Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
}

// Code goes to a Judge0 server for compile + run (ids are Judge0 CE language ids).
// Swap JUDGE0_URL for a self-hosted instance if the public one is unavailable.
// HTML, CSS, JSON and plain text have no runner and are insert-only.
const JUDGE0_URL = "https://ce.judge0.com";
const RUNNERS = {
  java: { id: 62 }, // OpenJDK; source is saved as Main.java
  javascript: { id: 63 }, python: { id: 71 }, csharp: { id: 51 },
  cpp: { id: 54 }, c: { id: 50 }, rust: { id: 73 }, r: { id: 80 }, sql: { id: 82 }, // SQLite
};

async function runCode() {
  const lang = $("lang").value;
  const runner = RUNNERS[lang];
  const out = $("output");
  out.hidden = false;
  out.className = "";
  if (!runner) { out.className = "error"; out.textContent = `Running ${LANGS[lang]} can't be run here, but you can still insert it.`; return; }
  const code = editor.getValue();
  if (!code.trim()) { out.textContent = "Nothing to run."; return; }

  $("run").disabled = true;
  out.textContent = "Running…";
  try {
    const res = await fetch(`${JUDGE0_URL}/submissions?base64_encoded=false&wait=true`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ source_code: code, language_id: runner.id }),
    });
    if (!res.ok) throw new Error(`Runner returned ${res.status}`);
    const r = await res.json();
    const failed = r.compile_output || r.stderr || r.message;
    out.className = failed ? "error" : "";
    out.textContent = (r.stdout || "") + (failed ? (r.stdout ? "\n" : "") + (r.compile_output || r.stderr || r.message) : "")
      || `(no output) — ${r.status?.description ?? ""}`;
    lastRun = { code, lang, text: out.textContent };
  } catch (e) {
    out.className = "error";
    out.textContent = "Could not run: " + e.message;
  } finally {
    $("run").disabled = false;
  }
}

function copyCode() {
  copyFormatted({ code: editor.getValue(), lang: $("lang").value, output: "" });
}

function initEditor(onCreated) {
  require.config({ paths: { vs: "https://cdn.jsdelivr.net/npm/monaco-editor@0.52.2/min/vs" } });
  require(["vs/editor/editor.main"], () => {
    editor = monaco.editor.create($("editor"), {
      value: placeholder(DEFAULT_LANG),
      language: DEFAULT_LANG, automaticLayout: true, minimap: { enabled: false },
      fontSize: 13, scrollBeyondLastLine: false,
      // the pane is narrow and fixed by OneNote, so trim the gutter to leave room for code
      glyphMargin: false, folding: false, lineDecorationsWidth: 4, lineNumbersMinChars: 2,
      wordWrap: "off", // long lines scroll sideways instead of wrapping
      scrollbar: { horizontal: "auto", alwaysConsumeMouseWheel: false },
      renderValidationDecorations: "off", // no squiggles, any language
    });
    $("lang").onchange = () => {
      const lang = $("lang").value;
      monaco.editor.setModelLanguage(editor.getModel(), lang);
      // swap the starter comment only if the user hasn't typed anything of their own
      const v = editor.getValue();
      if (!v.trim() || Object.keys(LANGS).some((l) => v === placeholder(l))) editor.setValue(placeholder(lang));
    };
    $("theme").onchange = () => monaco.editor.setTheme($("theme").value);
    if (onCreated) onCreated();
  });
}

// True inside OneNote; false when editor.html is opened as a normal web page.
let IN_OFFICE = false;

function start(info) {
  IN_OFFICE = !!(info && info.host);
  if (!IN_OFFICE) {
    document.body.classList.add("web");
    $("insert").textContent = "Copy for OneNote";
    $("insert").title = "Copy the formatted code block, then paste it into OneNote";
  }
  for (const [k, v] of Object.entries(LANGS)) $("lang").add(new Option(v, k));
  $("lang").value = DEFAULT_LANG;
  $("run").onclick = runCode;
  $("insert").onclick = insertIntoPage;
  $("copy").onclick = copyCode;
  $("expand").onclick = popOut;
  initEditor(IS_DIALOG ? initDialog : null);
}

// office.js resolves onReady with host = null outside Office; if it failed to load at all, start anyway.
if (window.Office) Office.onReady(start); else start(null);
