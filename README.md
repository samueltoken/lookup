<img src="docs/assets/logo.png" width="160" alt="lookup duck logo">

# lookup

lookup is a Windows document viewer for PDFs, Korean Hangul documents, and common Office files. Page navigation, search, printing, and annotations share the same reading window.

PDFs open directly. Other documents are converted on your computer and displayed as pages. Highlights, notes, and page edits are saved as PDFs; changes to the original document's text, formulas, or slide content belong in its editor.

[Download the latest release](https://github.com/samueltoken/lookup/releases/latest) · [Website](https://samueltoken.github.io/lookup/) · [Issues](https://github.com/samueltoken/lookup/issues)

## Installation

The desktop build targets **Windows x64**.

1. Open the [release page](https://github.com/samueltoken/lookup/releases/latest) and download `lookup-Setup-<version>.exe`.
2. Close lookup if it is running, then run the installer.
3. Launch lookup from the desktop or Start menu. Choose **Open**, or drag a document into the window.

To open files from File Explorer, choose **Open with → lookup**. Windows controls the default app for each file type; installing lookup does not guarantee that every existing default will change.

**Updating:** run the newer Setup installer. There is no need to uninstall first. The installer removes verified lookup files and shortcuts left by earlier installations while preserving documents and saved settings. Cleanup of a shared installation may require administrator permission.

**Removing lookup:** use **Windows Settings → Apps → Installed apps → lookup → Uninstall**. Deleting the downloaded Setup file only removes the installer, not the installed app.

## Supported files

| Format | Extensions |
| --- | --- |
| PDF | `.pdf` |
| Hangul | `.hwp`, `.hwpx` |
| Word | `.doc`, `.docx` |
| Excel | `.xls`, `.xlsx`, `.xlsm` |
| PowerPoint | `.ppt`, `.pptx` |
| Access | `.accdb` |
| HTML | `.html`, `.htm` |

Opening a non-PDF file creates a local copy for viewing. The conversion does not rewrite the source document.

Rendering depends on the file. Complex tables, embedded objects, uncommon fonts, and older Office or Hangul variants may look different from their original editor. lookup uses local conversion engines when available, with built-in readers as fallbacks. Installed Microsoft Office or LibreOffice may improve results for supported formats.

Excel documents include a sheet selector. Access support displays database tables; it does not reproduce forms or reports. lookup does not run spreadsheet macros to produce its reading view.

## Reading and saving

The **Document** tab holds page navigation and zoom. The **View** tab holds layouts, thumbnails, search, fullscreen, and the light/dark switch. **Edit** contains highlights, pen marks, text notes, rotation, and page deletion.

Use single-page, two-page, or current-page view to suit the document. When reading a zoomed page, the mouse wheel moves within that page. **Previous**, **Next**, `Page Up`, and `Page Down` move between pages.

Search highlights matching text and lists results by page. Image-only pages use text recognition, so a scan may take longer to search and recognition may miss some words.

**Save As** writes the displayed document and its edits to a PDF. Use it when the original should stay untouched. **Overwrite** updates an opened PDF after confirmation; converted files require **Save As**. Neither action writes annotations back into a Word, Excel, PowerPoint, or Hangul file.

### Keyboard shortcuts

| Shortcut | Action |
| --- | --- |
| `Ctrl+O` | Open a document |
| `Ctrl+F` | Search the document |
| `Page Up` / `Page Down` | Previous / next page |
| `Ctrl+=` / `Ctrl+-` | Zoom in / out |
| `Ctrl+0` | Reset zoom |
| `F11` | Enter / leave fullscreen |
| `Ctrl+D` | Switch light / dark mode |
| `Ctrl+S` | Overwrite the current PDF |
| `Ctrl+Shift+S` | Save As |
| `Ctrl+P` | Print |
| `Ctrl+Z` / `Ctrl+Y` | Undo / redo edits |

The interface supports English and Korean, under **Settings → Language**.

A document path also works from the command line:

```powershell
lookup.exe ".\example.pdf"
```

## Webpage capture

Choose **Webpage** and enter a web address to open a PDF capture in the reader. Save the result with **Save As**.

Capture depends on what the website sends to the app. Login screens, bot checks, interactive content, and images loaded on demand may prevent a complete result. If the result is incomplete, check the page in a browser before treating the PDF as a copy of the original.

## Local processing and internet access

There is no lookup account or document upload service. Local files are read and converted on your computer. Converted copies and preferences are stored in the app's local data.

Some operations still use the internet:

- Webpage capture contacts the address you entered and that page's resources.
- Update checks and downloads contact GitHub.
- Text recognition for scanned pages may download its runtime and language data when first needed. Recognition itself runs locally.

Documents are not sent to a remote conversion service. For public bug reports, remove confidential content from screenshots and sample files before sharing them.

## Development

Use a Windows machine with Git, Node.js, and npm installed.

```powershell
git clone https://github.com/samueltoken/lookup.git
cd lookup
npm ci
npm run setup:accdb
npm start
```

`setup:accdb` downloads and prepares the Java runtime used to read Access tables. It is needed for builds that include Access support. Microsoft Office and LibreOffice are separate, optional local applications, not npm dependencies.

Package the Windows installer or portable executable with:

```powershell
npm run dist
npm run dist:portable
```

Both commands write to `release/` and do not publish a release. The desktop app uses Electron and PDF.js. Conversion code lives in [`src/main/`](src/main/); the reading interface lives in [`src/renderer/`](src/renderer/).

The [project website](https://samueltoken.github.io/lookup/) is plain HTML and CSS in [`docs/`](docs/). It has no build step or external font dependency.

## Support

For bugs, [open an issue](https://github.com/samueltoken/lookup/issues). Include the lookup version, file format, steps to reproduce the problem, and the result you expected. A small sample that reproduces a rendering problem is more useful than a large private document.

For questions or private reports, email [lamsaiku65@gmail.com](mailto:lamsaiku65@gmail.com).

## License

MIT, as declared in [`package.json`](package.json). Third-party components retain their own licenses.
