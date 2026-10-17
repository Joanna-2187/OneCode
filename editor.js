const LANGS = {
  javascript: "JavaScript", typescript: "TypeScript", python: "Python", java: "Java",
  csharp: "C#", cpp: "C++", c: "C", go: "Go", rust: "Rust", html: "HTML", css: "CSS",
  sql: "SQL", json: "JSON", bash: "Shell", powershell: "PowerShell", plaintext: "Plain text",
};
// highlight.js name where it differs from Monaco's
const HLJS_NAME = { bash: "bash", powershell: "powershell", html: "xml", plaintext: "plaintext" };

// OneNote strips CSS classes, so token colours must be inlined.
const TOKEN_COLORS = {
  "hljs-keyword": "#0000ff", "hljs-built_in": "#267f99", "hljs-type": "#267f99",
  "hljs-literal": "#0000ff", "hljs-number": "#098658", "hljs-string": "#a31515",
  "hljs-regexp": "#811f3f", "hljs-comment": "#008000", "hljs-title": "#795e26",
  "hljs-function": "#795e26", "hljs-attr": "#e50000", "hljs-name": "#800000",
  "hljs-tag": "#800000", "hljs-variable": "#001080", "hljs-params": "#001080",
  "hljs-meta": "#af00db", "hljs-selector-class": "#800000", "hljs-property": "#001080",
};

let editor;
const $ = (id) => document.getElementById(id);

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
  const pre = "margin:0;font-family:Consolas,'Courier New',monospace;font-size:10pt";
  // OneNote collapses whitespace even inside <pre>, so use divs with <br> and &nbsp;.
  const outHtml = output
    ? `<div style="${pre};margin-top:6px;padding-top:6px;border-top:1px solid #ddd;color:#444">${keepWhitespace(escapeHtml(output))}</div>`
    : "";
  return `<table style="border-collapse:collapse"><tr><td style="background:#f5f5f5;border:1px solid #ddd;padding:6px 10px">` +
    `<div style="${pre};color:#1f1f1f">${keepWhitespace(holder.innerHTML)}</div>${outHtml}` +
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
  const html = buildHtml(code, lang, ran ? lastRun.text : "");
  Office.context.document.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, (r) => {
    setStatus(r.status === Office.AsyncResultStatus.Succeeded ? "Inserted." : "Failed: " + r.error.message);
  });
}

// Java can't run in the browser, so code goes to a Judge0 server for compile + run.
// Swap JUDGE0_URL for a self-hosted instance if the public one is unavailable.
const JUDGE0_URL = "https://ce.judge0.com";
const RUNNERS = { java: { id: 62 } }; // 62 = Java (OpenJDK); source is saved as Main.java

async function runCode() {
  const lang = $("lang").value;
  const runner = RUNNERS[lang];
  const out = $("output");
  out.hidden = false;
  out.className = "";
  if (!runner) { out.className = "error"; out.textContent = `Running ${LANGS[lang]} isn't supported yet (Java only).`; return; }
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

async function copyCode() {
  try { await navigator.clipboard.writeText(editor.getValue()); setStatus("Copied."); }
  catch { setStatus("Copy blocked."); }
}

function initEditor() {
  require.config({ paths: { vs: "https://cdn.jsdelivr.net/npm/monaco-editor@0.52.2/min/vs" } });
  require(["vs/editor/editor.main"], () => {
    editor = monaco.editor.create($("editor"), {
      value: "// Write code here, then click Insert into page\n",
      language: "javascript", automaticLayout: true, minimap: { enabled: false },
      fontSize: 13, scrollBeyondLastLine: false,
      renderValidationDecorations: "off", // no squiggles, any language
    });
    $("lang").onchange = () => monaco.editor.setModelLanguage(editor.getModel(), $("lang").value);
    $("theme").onchange = () => monaco.editor.setTheme($("theme").value);
  });
}

Office.onReady(() => {
  for (const [k, v] of Object.entries(LANGS)) $("lang").add(new Option(v, k));
  const setHeight = (h) => {
    h = Math.max(100, Math.min(900, Math.round(h)));
    $("wrap").style.flex = "none";
    $("wrap").style.height = h + "px";
    $("size").value = h;
  };
  const setWidth = (w) => {
    w = Math.max(200, Math.round(w));
    $("wrap").style.alignSelf = "flex-start"; // stop the column layout stretching it
    $("wrap").style.width = w + "px";
  };
  $("size").oninput = () => setHeight(+$("size").value);
  // Hover the bar on the editor's right edge (left-right arrows), then click and drag.
  $("resizer").onpointerdown = (e) => {
    e.preventDefault();
    const startX = e.clientX, startW = $("wrap").offsetWidth;
    $("resizer").setPointerCapture(e.pointerId);
    $("resizer").onpointermove = (m) => setWidth(startW + m.clientX - startX);
    $("resizer").onpointerup = () => { $("resizer").onpointermove = $("resizer").onpointerup = null; };
  };
  $("run").onclick = runCode;
  $("insert").onclick = insertIntoPage;
  $("copy").onclick = copyCode;
  initEditor();
});
