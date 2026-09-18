# PDF Page Editor: Windows Context Menu (Personal Windows Utility)

A local Windows-only PDF creation/manipulation utility focused on handling files from the context menu in file explorer.

- tray startup (always available)
- File Explorer right-click action for converting images to PDF
- File Explorer right-click action for combining and splitting PDFs + images
- visual page list with split starts
- per-page rotation in 90-degree steps
- page reordering with safe replace-original or Save As output
- editable output names with defaults

## 1) Install

From this folder in PowerShell:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
./scripts/install.ps1
```

This will:

- install dependencies from `requirements.txt`
- register `Split PDF...` in your PDF right-click menu (current user)
- register `Combine to PDF...` for PDFs and common image files (current user)
- register `Convert to PDF...` for common image files (current user)
- create a startup shortcut so the tray app launches at logon
- launch the tray app now

Important:

- The install registers the current cloned checkout as the live app location.
- If you move, rename, or delete this project folder later, Explorer integration and tray startup will break until you run `./scripts/uninstall.ps1` and then `./scripts/install.ps1` from the new location.

## 2) Use

- Right-click any PDF in File Explorer -> `Split PDF...`
- In the app:
  - move or rotate pages, then use `Save Changes` to replace the original safely
  - turn off `Replace original` to use `Save As...` with an `_Edited.pdf` default name
  - check `Start split here` on pages that begin a new output file
  - edit section names on the right
  - optionally rotate selected pages with `Rotate -90` / `Rotate +90`
  - click `Export Splits`
- Output files are written next to the source PDF.
- Existing names are not overwritten; `_1`, `_2`, etc. are appended.

- Multi-select PDFs in File Explorer -> `Combine to PDF...`
- The combine flow opens only a standard Windows save dialog.
- The save dialog defaults to the source folder and suggests `combined.pdf`.
- Supported image types: `.jpg`, `.jpeg`, `.png`, `.bmp`, `.tif`, `.tiff`, `.webp`, `.heic`, `.heif`
- Images are compressed with a balanced profile before being written into the merged PDF.
- On success, the merged PDF opens automatically and the original selected source files are deleted.

- Right-click one or more images in File Explorer -> `Convert to PDF...`
- The convert flow uses the same save dialog, compression, and progress window.
- On success, the PDF opens automatically and the original selected image files are deleted.

## 3) CST workflow (date + facility)

Right-click the tray icon and choose **Open CST PDF (date + facilities)…**.

1. Set the document date at the top of the right panel (it applies to every section).
   The ◀ ▶ buttons step one day at a time.
2. Click page images to mark the beginning of each section. The new section's
   facility field takes focus so you can type right away.
3. Type to search a facility for every section. **Tab** accepts the first match (or
   the highlighted one) and moves to the next field. The list is downloaded from the
   same deployed `Facilities.txt` the RMRR app uses each time a PDF opens, so every
   choice is a name the form accepts. The last download is used when offline.
   Each section is filed as **Tracking** unless you switch its selector to **Delivery**.
4. **Export Splits** writes a batch folder of dated PDFs plus `batch.json`.
   **Export & submit to RMRR** also sends each section through the RMRR app's
   `/tracking/v2` API with the selected technician, date, facility, and type.

The export folder is asked for once and remembered. Once the server confirms every
section, the window closes, a tray notification reports the result, and the local
copies are deleted: the batch folder, plus the downloaded PDF when it came from the
Outlook intake (the email itself is untouched). A PDF you opened by hand is kept.

`submission.json` in the batch folder stores each section's submission ID and receipt.
After an interruption, use **Submit / retry saved CST batch** with that folder's
`batch.json`: received sections are skipped, and the server ignores a repeated ID, so
nothing is uploaded twice. A batch that did not finish keeps all its files. PDFs over
the API's 10 MB limit are rejected before sending.
Success means received by RMRR for processing; the backend owns SharePoint archival.

### Optional automatic Outlook intake

Turn on **Automatically open emailed CST PDFs** in the tray menu (off by default). It needs
Outlook Classic running with the office mailbox, and `pywin32` from `requirements.txt`.

- Every 60 seconds it checks the office inbox for PDF attachments from the senders
  in `SENDERS` (`app/outlook_intake.py`: James and Ryan) received since the start of the day intake was
  first enabled (at most the last 14 days). Each attachment is downloaded once and queued under
  `%LOCALAPPDATA%/PDFSplitter/CST Intake`, so it catches up after a restart or pause.
- Queued PDFs open one at a time in the CST workspace and never replace a document
  you are working on. After an export, the next one opens. The technician is set
  from the sender; for a PDF opened by hand, pick it under the date.
- Any email text beyond the recognized mobile signature is shown when the PDF opens.
- **Set aside current queued PDF** drops an unwanted or unreadable file from the queue.
- Once every section of an emailed PDF is confirmed received, the email moves from the
  Inbox to that month's CST folder (`0926` > `CST 0926`, by the email's received date).
  If the folder does not exist yet, the email stays in the Inbox and the result says so.
- Unexported split choices are not saved across a restart.

The status line in the tray menu reports watching, waiting for Outlook, or an error.

## 4) Optional CLI

From this project root:

```powershell
python -m app.tray_runtime
python -m app.launcher "C:\path\to\file.pdf"
python -m app.launcher combine "C:\path\to\file1.pdf" "C:\path\to\image.png"
python -m app.launcher convert-image "C:\path\to\image.png"
python -m app.launcher convert-image "C:\path\to\image1.png" "C:\path\to\image2.jpg"
```

The same launcher command can be registered as a SumatraPDF external viewer so the current PDF opens directly in the page editor.

## 5) Uninstall

```powershell
./scripts/uninstall.ps1
```

This removes startup + right-click integration.

## Troubleshooting

- If right-click does nothing, restart the tray app:
  1. Right-click tray icon -> `Quit`
  2. Run `pythonw .\\run_tray.py` from project root
- Check logs at:
  - `%LOCALAPPDATA%\\PDFSplitter\\app.log`
