# Office PDF Binder

English | [日本語](README.ja.md)

Office PDF Binder is a Windows desktop application that combines PDF, Word,
Excel, and PowerPoint files into a single PDF. You can organize, delete, and
rotate individual PDF pages before saving the result.

![Office PDF Binder main window](docs/images/screenshot-main-en.png)

---

## 1. Requirements

- Windows 10 or Windows 11 (64-bit)
- Microsoft Office when converting Word, Excel, or PowerPoint files
- No separate Python installation is required for packaged releases

Office files are converted through the locally installed Microsoft Office
applications. Office PDF Binder does not upload documents to an online service.
PDF files are loaded as individual pages. Word, Excel, and PowerPoint documents
are added as files and converted to PDF when the combined PDF is saved.

---

## 2. Features

- Load, reorder, delete, and rotate individual PDF pages
- Add Word, Excel, and PowerPoint documents as files and convert them when saving
- Add files with a file dialog, drag and drop, or Windows Explorer
- Reorder, delete, and rotate PDF pages
- Reorder multiple selected pages by dragging them together
- Detect duplicate files
- Import existing PDF bookmarks
- Create automatic bookmarks for each source file
- Add, rename, delete, and navigate manual bookmarks
- Add headers, footers, dates, and page numbers
- Export selected pages as a PDF
- Export selected PDF pages as JPEG images at 300 dpi
- Open the original file by double-clicking a page
- Five thumbnail zoom levels with Ctrl+mouse wheel support
- Undo and redo up to 15 changes

Supported extensions:

`.pdf / .docx / .doc / .docm / .xlsx / .xls / .xlsm / .pptx / .ppt / .pptm`

---

## 3. Installation and Portable Version

Download the latest installer or portable ZIP from GitHub Releases.

The installer includes the application, this manual, license documents, and
the corresponding source archive. The installer can be displayed in English
or Japanese, and Office PDF Binder uses the language selected during setup.

The portable version automatically uses Japanese on a Japanese Windows system
and English on other Windows language settings.

### Installation notes

- Uninstall an older release before installing a version that cannot be
  upgraded in place.
- This is an unsigned independently developed application. Windows SmartScreen
  may display a warning.
- Microsoft Office is required only for Word, Excel, and PowerPoint conversion.

---

## 4. Basic Use

### 4.1 Add files

- Select **Add Files**, press `Ctrl+O`, or drag supported files into the window.
- You can select multiple supported files in Windows Explorer and use
  **Open with Office PDF Binder**.
- Files already present in the list are detected and are not added twice.

### 4.2 Organize pages

- Select one or more pages to enable page operations.
- Rotate pages 90 degrees left or right.
- Move pages up, down, to the top, or to the bottom.
- Press `Delete` to remove selected items and `Ctrl+A` to select all items.
- Drag selected pages to reorder them while preserving their relative order.

Office documents appear as a single item in the list. They are converted to
PDF when the combined document is saved.

### 4.3 Start a new session

Select **File > New** to clear the current list, bookmarks, and page settings.

### 4.4 Bookmarks

- Select **View > Bookmarks** to show or hide the bookmarks panel.
- Existing PDF bookmarks are imported when a PDF is added.
- Enable **Settings > Create Bookmarks for Each File** to create a bookmark at
  the first page of each source file.
- Right-click a selected page and choose **Add Bookmark to Selected Page**.
- Use the bookmarks panel to add, rename, or delete bookmarks.
- Double-click a bookmark to navigate to its page.
- Renaming an automatic bookmark converts it to a manual bookmark.

### 4.5 Zoom

- Use **View > Zoom In**, **Zoom Out**, or **Fit to Window**.
- Hold `Ctrl` while using the mouse wheel to zoom.
- Press `Ctrl+0` to fit the thumbnails to the window.

### 4.6 Headers, footers, and page numbers

Select **Settings > Header and Footer** to configure content added when the PDF
is saved.

- Enable the header, footer, or both.
- Enter separate text for the left, center, and right positions.
- Insert the current date on the right.
- Select the page-number position and format.
- Set the first page number and font size.

![Header and footer settings](docs/images/screenshot-header-footer-en.png)

### 4.7 Export selected pages

- Select **File > Export Selected Pages as PDF** to create a PDF containing
  only the selected pages.
- Select **File > Export Selected Pages as Images** to export selected PDF pages
  as JPEG images at 300 dpi.
- Word, Excel, and PowerPoint items cannot be exported as images.

### 4.8 Combine and save

Select **File > Save As** or press `Ctrl+S`. Choose an output path, and Office
PDF Binder combines the current list into one PDF.

After saving, the PDF opens in the default PDF viewer. The application then
asks whether to clear the current list.

---

## 5. Keyboard Shortcuts

| Action | Shortcut |
|---|---|
| New | `Ctrl+N` |
| Add Files | `Ctrl+O` |
| Combine and Save | `Ctrl+S` |
| Exit | `Ctrl+Q` |
| Select All | `Ctrl+A` |
| Delete Selected | `Delete` |
| Zoom In / Out | `Ctrl` + `+` / `-`, or `Ctrl` + mouse wheel |
| Fit to Window | `Ctrl+0` |
| Header and Footer | `Ctrl+H` |
| Undo / Redo | `Ctrl+Z` / `Ctrl+Y` |

---

## 6. Troubleshooting

### An Office file cannot be converted

- Verify that the corresponding Microsoft Office application is installed.
- Close the document in other applications and try again.
- Password-protected or damaged documents may not be convertible.

### A PDF cannot be saved

- Check that you have write permission for the output folder.
- Keep at least 100 MB of free space on the output drive.
- Close the output PDF if it is open in another application.

### Pages cannot be exported as images

Image export supports PDF pages only. Word, Excel, and PowerPoint items are not
included.

### High memory use or slow thumbnails

- Split very large documents into smaller operations when practical.
- Reduce the thumbnail zoom level.

---

## 7. Restricted Portable Mode

The portable package contains `OfficePDFBinder.restricted-portable`. Keep this
marker file next to the executable.

In restricted portable mode, Office PDF Binder:

- does not save application settings in AppData;
- does not write its optional debug log to the user profile; and
- creates temporary Office-conversion PDFs beside the source Office file and
  deletes them after processing.

Microsoft Office, Windows, and Qt may still use their own system services,
temporary folders, caches, or registry settings.

---

## 8. Development and Testing

Install runtime and development dependencies with:

```powershell
pip install -r requirements.txt
pip install -r requirements-dev.txt
```

Run automated tests with:

```powershell
python -m pytest
```

Build operations use the single `build.ps1` entry point. To regenerate only
the manual and installer from the existing `dist`, run:

```powershell
.\build.ps1 -Mode Package
```

After application code changes, reuse Nuitka intermediate files with:

```powershell
.\build.ps1 -Mode Fast
```

For the final public release, run a clean build:

```powershell
.\build.ps1 -Mode Release
```

`Fast` and `Release` create the distribution directory, portable package, and
installer. `Package` skips Nuitka and regenerates only the installer from the
existing distribution directory.

The Nuitka application is compiled once. The same distribution directory is
used for the installer and the marker-based portable ZIP.

---

## 9. License

- Office PDF Binder: GNU Affero General Public License v3.0 (`LICENSE.txt`)
- Third-party components: see `NOTICE.txt`
- Corresponding source: included as `source.zip` in installer distributions

---

Copyright (C) 2026 Takeshi Kashiwagi
