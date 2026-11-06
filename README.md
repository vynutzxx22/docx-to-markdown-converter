# DOCX to Markdown Converter

A Windows desktop utility that converts Microsoft Word `.docx` documents into Markdown (`.md`). It supports batch conversion, dark and light themes, Markdown tables, headings, lists, hyperlinks, and image extraction.

> **Project status:** Initial public release preparation. The Windows executable is built automatically by GitHub Actions when a version tag is pushed.

## Features

- Convert one DOCX file or a batch of files.
- Convert Word title and heading styles into Markdown headings.
- Convert Word tables into GitHub Flavored Markdown tables.
- Preserve external hyperlinks as Markdown links.
- Convert numbered and bulleted paragraphs into Markdown list syntax.
- Extract embedded images into a companion media directory.
- Select an output directory.
- Switch between dark and light themes.
- View conversion progress and per-file errors.

## Download

Open the repository's **Releases** page and download `DOCX-to-Markdown-Converter.exe` from the latest release.
Or
You can download the latest standalone binaries and source code directly without opening the Releases page:

| Asset | Platform / Type | Direct Download Link |
| :--- | :--- | :--- |
| 📦 **DOCX-to-Markdown-Converter.exe** | Windows x64 (Standalone Executable) | [**Download Latest .exe**](https://github.com/vynutzxx22/docx-to-markdown-converter/releases/latest/download/DOCX-to-Markdown-Converter.exe) |
| 🗜️ **Source code (zip)** | Archive | [**Download .zip**](https://github.com/vynutzxx22/docx-to-markdown-converter/archive/refs/tags/v1.0.0.zip) |
| 🗜️ **Source code (tar.gz)** | Archive | [**Download .tar.gz**](https://github.com/vynutzxx22/docx-to-markdown-converter/archive/refs/tags/v1.0.0.tar.gz) |

> ℹ️ **Note on Windows SmartScreen:**  
> The executable is a single-file PyInstaller build for Windows x64 and does not require Python to be installed. Because it is not digitally signed with a paid certificate, Windows SmartScreen might show a security warning. If prompted, click **"More info"** and then select **"Run anyway"**.

---

## Run from source

Requirements: Windows 10/11 and Python 3.10 or newer.

```powershell
py -m pip install -r requirements.txt
py app.py
```

## Build the portable EXE locally

1. Install Python 3.10 or newer and enable **Add Python to PATH**.
2. Extract or clone this repository.
3. Run `build_windows.bat`.
4. Find the executable at `dist/DOCX-to-Markdown-Converter.exe`.

Or run the build commands manually:

```powershell
py -m pip install -r requirements.txt
py -m PyInstaller --noconfirm --clean --onefile --windowed --name "DOCX-to-Markdown-Converter" app.py
```

## Usage

1. Launch the application.
2. Select one or more `.docx` files, or add a folder.
3. Choose an output folder. If you leave it blank, the app creates `Markdown_Output` beside the first source document.
4. Keep **Extract images** enabled if the document contains pictures.
5. Click **Convert documents**.

For each input file, the app creates a `.md` file. If image extraction is enabled, it also creates a matching `<document>_media` folder. Keep that folder beside the Markdown file when moving or sharing the output.

## Conversion notes

Markdown cannot represent every Word layout feature. Merged cells, nested tables, floating images, text boxes, headers and footers, footnotes, tracked changes, and complex numbering may need manual review. The converter prioritizes document text and common structural elements. Always review converted files before publication.

The application currently supports `.docx`. Open legacy `.doc` files in Microsoft Word and save them as `.docx` first.

## Development and tests

```powershell
py -m pip install -r requirements-dev.txt
py -m pytest
```

## Release process

1. Update the version in `app.py` and `CHANGELOG.md`.
2. Commit and push changes to `main`.
3. Create and push a version tag, for example:

```powershell
git tag v1.0.0
git push origin v1.0.0
```

4. GitHub Actions builds the Windows executable and attaches it to a GitHub Release.

## License

This project is distributed under the MIT License. See [LICENSE](LICENSE).
