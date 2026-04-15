# ENEX → OneNote One-Click Helper

This is a browser-based tool that converts Evernote export files (`.enex`) into Word documents (`.docx`) that OneNote can ingest quickly.

## Why this is the easiest practical flow

Direct ENEX → native OneNote package conversion is not officially supported by Microsoft. The lowest-friction workflow is:

1. Convert each Evernote note to `.docx`.
2. Drag DOCX files into a OneNote section.
3. OneNote creates one page per document.

This tool automates step 1 with batch handling and structure-aware output folders.

## Features

- Drag-and-drop or file picker upload for one or more `.enex` files.
- Progress bar and status updates during conversion.
- Error panel that reports notes/files that failed.
- Structured ZIP output (one folder per ENEX notebook) to preserve notebook grouping.
- Optional direct single-file download when exactly one note is converted.

## How to use

1. Open `index.html` in a modern browser (Chrome/Edge/Firefox).
2. Drop one or more `.enex` files or select them in the file picker.
3. (Optional) Enable **Download a single .docx when exactly 1 note is converted**.
4. Click **Convert & Download**.
5. If a ZIP is downloaded, extract it.
6. Open OneNote desktop and drag generated `.docx` file(s) into the destination section.

## What gets preserved

- Note title
- Text formatting (most common ENML formatting)
- Tables and checklists (best effort)
- Links
- Inline images (mapped using ENEX resource hashes)
- Notebook-level grouping (as ZIP folders named after each ENEX file)

## Limitations

- Native `.one` file output is not supported.
- Some advanced Evernote-only elements may be flattened.
- Non-image attachments are represented as placeholders in the output document.
- Very large exports may take time in-browser.

## Files

- `index.html` – UI and usage instructions
- `styles.css` – styling
- `app.js` – ENEX parsing + DOCX conversion + download logic
