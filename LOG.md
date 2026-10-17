## 2026-10-14 — Run Java
Goal: run snippets from the pane, Java first.

Finding: OneNote desktop can't host Office.js add-ins (web only), so the pane can't run code there.
Decision: v1 is a standalone Python tool over OneNote COM ([onecode_run.py](onecode_run.py)), not a plug-in.
Ribbon button is a stretch goal.

How a run works: write `#!java` as the first line of a block, press Ctrl+Alt+R (or run the script).
The tool reads the current page XML, runs each marked block (10s timeout), and writes the output in a
monospace block under the code, replacing the previous output (tagged with meta `onecode.output`).

## 2026-10-17 — Restored add-in files
Restored the earlier task-pane add-in into [mainfest.xml](mainfest.xml), [editor.html](editor.html) and
[editor.js](editor.js) ([editor.css](editor.css) was already there). Pane files are named `editor.*`
(the old `taskpane.*` names are gone); manifest URLs now point to `https://joanna-2187.github.io/OneCode/editor.html` (repo is Joanna-2187/OneCode; manifest renamed to `manifest.xml`).
The pane's Run button posts Java to the public Judge0 server (web only). Nothing is on GitHub yet, so
the manifest's Pages URLs don't resolve until it's published. Nothing here has been run.

Status of onecode_run.py: written, not yet tested against a real OneNote page. Needs `pip install pywin32` and `java` 11+ on PATH.
