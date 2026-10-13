<img src="docs/assets/logo.png" width="88" alt="lookup duck logo">

# lookup

A document viewer for Windows. Open a file, find what matters, and get back to work.

[Download for Windows](https://github.com/samueltoken/lookup/releases/latest) · [Website](https://samueltoken.github.io/lookup/) · [Report an issue](https://github.com/samueltoken/lookup/issues)

lookup brings PDFs, Korean HWP documents, and Office files into one reading window. It includes text search, page thumbnails, zoom, fullscreen, printing, highlights, pen marks, and text notes. There is no account to create or document upload service.

## Install

Download the **Setup** installer from the [latest release](https://github.com/samueltoken/lookup/releases/latest), close any running instance of lookup, and run the installer. Windows x64 is the build target.

The installer cleans up verified files and shortcuts from older lookup installations. Documents and saved settings are kept; cleanup in shared Windows folders may request administrator permission.

Open documents with **Open**, drag a file into the window, or use **Open with → lookup** in File Explorer. To remove the app, use **Windows Settings → Apps → Installed apps → lookup → Uninstall**. Deleting the downloaded installer does not remove the installed app.

## File formats

| Type | Extensions |
| --- | --- |
| PDF | `.pdf` |
| Hangul | `.hwp`, `.hwpx` |
| Word | `.doc`, `.docx` |
| Excel | `.xls`, `.xlsx`, `.xlsm` |
| PowerPoint | `.ppt`, `.pptx` |
| Access tables | `.accdb` |
| HTML | `.html`, `.htm` |

PDFs open directly. Other formats are converted locally for viewing; the source file is not rewritten during conversion. Saving annotations or page changes is a separate action, so use **Save As** to keep an untouched original.

Conversion is not identical to opening a file in its original editor. Complex layouts, embedded objects, unusual fonts, and some older files may render differently. Installed Microsoft Office or LibreOffice may improve conversion for supported formats. Access support reads tables, not forms or reports.

The **Webpage** command opens a web address as a PDF. Login requirements, bot checks, and content that loads on demand may prevent a complete capture.

## Everyday controls

| Shortcut | Action |
| --- | --- |
| `Ctrl+O` | Open a file |
| `Ctrl+F` | Find text |
| `Ctrl+Shift+S` | Save As |
| `Ctrl+P` | Print |
| `Ctrl+Z` / `Ctrl+Y` | Undo / redo edits |

A file path also works from the command line:

```powershell
lookup.exe "C:\Documents\example.pdf"
```

## Local processing

Local documents are rendered and converted on your computer. Opening a web address contacts that website, and update checks contact GitHub. lookup does not send documents to a remote conversion service.

## Build from source

On Windows, with Node.js and npm installed:

```powershell
npm ci
npm run setup:accdb
npm start
```

Create an installer or a portable build:

```powershell
npm run dist
npm run dist:portable
```

Builds are written to `release/`. `setup:accdb` downloads and prepares the separate Java runtime for Access table conversion. Microsoft Office and LibreOffice are optional local conversion engines, not npm dependencies.

The website is plain HTML and CSS in [`docs/`](docs/), with no build step.

## Bugs and feedback

Please [open an issue](https://github.com/samueltoken/lookup/issues) with the app version, file type, what happened, and what you expected. A screenshot or a small, non-sensitive sample helps with rendering problems. Keep private documents and personal details out of public issues.

For other questions, contact [lamsaiku65@gmail.com](mailto:lamsaiku65@gmail.com).

## License

MIT, as declared in [`package.json`](package.json). Bundled third-party components retain their own licenses.
