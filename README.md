EMLtoPDF Automation — End‑to‑End README
====================================

IMPORTANT!: If you updated the argument names in UiPath. Then you'll need to update the command files under folder 'VBA Command Files'. These files are used by UiPath to invoke the Shell/VBScripts. 

Overview
--------
This solution automates converting only *digitally signed* emails (EML files) to PDF. It uses two VB scripts and two UiPath workflows:

- **CheckEMLSignature.vbs** — Detects whether an .eml file carries an Outlook/S/MIME digital signature.
- **eml2pdf.vbs** — Converts a given .eml to a PDF (including core headers/body).
- **Workflow 1: Signature Filter** — Invokes CheckEMLSignature for each .eml in the Input folder, logs results, and produces a filtered list of only the signed emails. Then invokes Workflow 2.
- **Workflow 2: Signed EML → PDF** — Invokes eml2pdf for each signed email and saves the PDFs to the Output folder with full logging.

Folder Structure (recommended)
------------------------------
- `Input/` — Drop .eml files here.
- `Output/` — PDFs land here (mirrors EML base names).
- `Scripts/` — Contains `CheckEMLSignature.vbs` and `eml2pdf.vbs`.
- `Logs/` — UiPath run logs and any optional CSV lists.
- `Temp/` (optional) — Scratch space for intermediate artifacts.

Prerequisites
-------------
- Windows with **cscript.exe** (Windows Script Host).
- UiPath Studio/Robot (project name suggested: **EMLtoPDF**).
- Read/write access to the above folders.
- (Optional) Orchestrator, if scheduling/queuing is desired.

Configuration
-------------
Set these values in UiPath (Config file/Assets/Arguments—choose what fits your setup):
- `InputFolder` — path to the Input directory
- `OutputFolder` — path to the Output directory
- `ScriptsFolder` — path to the VB scripts
- `LogFolder` — path for logs (and optional CSV results)
- `ExecutionTimeoutSec` — max wait per script call
- `LogStdoutStderr` — true/false to capture script output verbatim

What the VB Scripts Do
----------------------
**CheckEMLSignature.vbs**
- Input: full path to an `.eml` file.
- Output (stdout): a concise status line (e.g., `SIGNED`, `UNSIGNED`, or a reason on error).
- Exit code: `0` = success, non‑zero = error.
- Purpose: Enable Workflow 1 to filter only S/MIME‑signed emails.

**eml2pdf.vbs**
- Input: full path to an `.eml` file and an output PDF path.
- Output (stdout): progress/diagnostics; (stderr): errors if any.
- Exit code: `0` = success, non‑zero = error.
- Purpose: Convert an .eml to PDF, including key headers/body (and typical inline/HTML content when present).

Workflow Details
----------------
**UiPath Workflow 1 — Signature Filter**
1. UiPath reads 2 paths (`InputFolder` --> conatains all EML files, and the path of the CheckEMLSignature.vbs script).
2. UiPath uses the paths to execute the CheckEMLSignature.vbs script:
   ```
   cscript //nologo "<ScriptsFolder>\CheckEMLSignature.vbs" "<InputFolder>"
   ```
3. Parse stdout. If `SIGNED`, add to an in‑memory list.
4. (Optional) Persist the list to `Logs/SignedList.csv` for traceability.
5. Log stdout, stderr, and exit codes for each file.
6. The digitally signed EML files are moved to the Filtered folder (Input\signed)
7. The EML files who are not digitally signed are moved to a sepreate folder (Input\Notsigned)
6. Invokes Workflow 2 

**UiPath Workflow 2 — Signed EML → PDF**
1. Receive the filtered list of signed `.eml` paths.
2. For each file, compute the destination:
   - `OutputPath = <OutputFolder>\<BaseName>.pdf`
3. Run to convert EML to PDF:
   ```
   cscript //nologo "<ScriptsFolder>\eml2pdf.vbs" "<FullPathToSignedEML>" "<OutputPath>"
   ```
4. Log stdout/stderr and exit code.
5. On success, verify file exists and is non‑empty; log “Converted” with the PDF path.
6. On failure, record the error; (optional) move the `.eml` to `Input/Failed/`.

Running the Automation
----------------------
**From UiPath Assistant (recommended):**
1. Run **EMLtoPDF — 1. Signature Filter**.
2. Then run **EMLtoPDF — 2. Convert Signed**.
   - If using a CSV handoff, ensure Workflow 2 reads `Logs/SignedList.csv`.
   - If using arguments/queue, ensure Workflow 1 publishes the list accordingly.

**Manual Script Testing (for troubleshooting):**
```
cscript //nologo "C:\Path\To\Scripts\CheckEMLSignature.vbs" "C:\Path\To\Input__EML_folder\"
cscript //nologo "C:\Path\To\Scripts\eml2pdf.vbs" "C:\Path\To\Input\Input__EML_folder\Signed\ "C:\Path\To\Output\"
```

Logging 
-----------------------
- Both workflows capture `stdout`, `stderr`, and exit codes for each script call.
- Logs are written to UiPath logs and (optionally) mirrored into `Logs/` files.
- Recommended log fields: Timestamp, Activity, File Path, Command, Exit Code, Stdout Hash (first 200 chars), Stderr Hash (first 200 chars).

Notes:
--------------
- S/MIME signature trust and validation depend on the Windows certificate store and local policy.
- Only digitally signed emails are converted by design (others are ignored).

FAQ
---
**Q: Can I convert unsigned emails too?**  
A: Yes—run `eml2pdf.vbs` directly or add a third workflow for unsigned items.

**Q: Where do I change Input/Output folders?**  
A: In UiPath Config/Assets/Arguments—whichever pattern you chose.

