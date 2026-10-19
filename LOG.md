## 2026-10-14 — Run Java
Goal: run snippets from the pane, Java first.
Finding: OneNote desktop can't host Office.js add-ins (web only), so the pane can't run code there.
Decision: v1 is a standalone Python tool over OneNote COM ([onecode_run.py](onecode_run.py)), not a plug-in.
How it works: `#!java` as a block's first line, then Ctrl+Alt+R. It runs each marked block (10s timeout) and
writes the output under the code, replacing the previous output (meta `onecode.output`).
Status: not yet tested against a real OneNote page. Needs `pip install pywin32` and `java` 11+ on PATH.

## 2026-10-17 — Restored add-in files
Restored the task-pane add-in: [manifest.xml](manifest.xml), [editor.html](editor.html), [editor.js](editor.js),
[editor.css](editor.css). Manifest URLs point to `https://joanna-2187.github.io/OneCode/editor.html`, which won't
resolve until the repo is published. Run posts Java to the public Judge0 server (web only). Not run.

## 2026-10-18 — Retro restyle
Styling only; not run or committed.
- Buttons: rounded (10px), retro palette (pink, blue, purple, green), raised 3px edge. Insert is purple, Copy dark
  blue with light-blue text, Run green.
- Dropdowns: 6px rounded, regular weight, raised edge. Light mode green, dark mode pink; open lists lighter and
  rounded to match (`appearance: base-select`, Chromium 135+ only).
- New **Retro** theme: CRT scanlines and flicker, VT323 font, built on the dark styling. The editor is black with
  neon green text, bright green comments and yellow strings; page chrome keeps the earlier purple look. Retro
  buttons use bright fills with black text.
- Hover notes (`title`) on every button and dropdown.
