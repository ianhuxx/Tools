# ENEX → OneNote One-Click Helper

This is a simple front-end tool to convert Evernote export files (`.enex`) into Word files (`.docx`) that OneNote can ingest quickly.

## Why this is the easiest practical flow

Direct ENEX → native OneNote package conversion is not officially supported by Microsoft. The lowest-friction workflow is:

1. Convert each Evernote note to `.docx`.
2. Drag all docs into a OneNote section.
3. OneNote creates one page per document.

This tool automates step 1 in one click.

## How to use

1. Open `index.html` in a modern browser (Chrome/Edge/Firefox).
2. Select one or more `.enex` files exported from Evernote.
3. Click **Convert & Download ZIP**.
4. Extract the ZIP.
5. Open OneNote desktop and drag all generated `.docx` files into the destination section.

## What gets preserved

- Note title
- Text formatting (most common ENML formatting)
- Tables and checklists (best effort)
- Links
- Inline images (mapped using ENEX resource hashes)

## Limitations

- Some advanced Evernote-only elements may be flattened.
- Non-image attachments are represented as placeholders in the output document.
- Very large exports may take time in-browser.

## Files

- `index.html` – UI and usage instructions
- `styles.css` – styling
- `app.js` – ENEX parsing + DOCX conversion + ZIP download
