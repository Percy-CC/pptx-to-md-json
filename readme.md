# PPTX to AI

Convert PowerPoint `.pptx` presentations into Markdown and JSON for searching, reviewing, or using as text input to AI tools. Conversion runs locally; this project does not send presentations to an AI provider.

## Features

- Batch-convert presentations from a folder or select files in the desktop app.
- Extract slide text, paragraph indentation, titles, hidden-slide status, and speaker notes.
- Preserve common text formatting in Markdown, including links and non-default text colors.
- Extract table contents, chart titles and series values, SmartArt text, image alternative text, and embedded-object labels.
- Produce both a readable Markdown corpus and a structured JSON file.

This tool extracts text and supported structured content; it does not render slides, perform OCR, or include image binaries in the output.

## Quick start: Windows app

The repository includes `PPTX_Converter.exe`. Run it, add one or more `.pptx` files (or add a folder), choose an output folder, and select **開始轉換** (Start Conversion). The interface and log messages are currently in Traditional Chinese.

## Run from Python

### Requirements

- Windows for the desktop GUI and included executable.
- Python installed and available as `py` (or substitute `python` in the commands below).
- `python-pptx` and `lxml`.

Install the dependencies:

```powershell
py -m pip install python-pptx lxml
```

### Desktop interface

```powershell
py convert_gui.py
```

Choose `.pptx` files or a folder, select the destination directory, and start conversion. If no output directory is selected, the app initially suggests an `output` folder beside the first selected presentation.

### Command-line converter

Place `.pptx` files directly inside the project's `input` folder, then run:

```powershell
py convert_pptx.py
```

The converter writes its results to the project's `output` folder. It processes files in sorted order and does not search subfolders.

## Generated files

Each conversion creates or replaces these files in the chosen output directory:

- `all_slides.md` — a readable corpus grouped by presentation and slide. Hidden slides are marked, and speaker notes are included when available.
- `all_slides.json` — an array of slide records containing the source filename, slide number, hidden status, title, content blocks, and speaker notes.

The JSON `blocks` array can contain:

- `text` blocks with paragraphs, levels, plain text, and Markdown text.
- `table` blocks with rows and Markdown table output.
- `chart` blocks with chart title and extracted series data in Markdown table form.
- `smartart` blocks with text and nesting levels.
- `image` blocks containing image alternative text, when present.
- `object` blocks identifying embedded or linked OLE objects.

## Project files

- `convert_pptx.py` — extraction and output generation.
- `convert_gui.py` — Tkinter desktop application.
- `PPTX_Converter.spec` — PyInstaller configuration for building the Windows app.
- `project.md` — project purpose, scope, components, and output contract.
- `input/` and `output/` — default input and generated-output directories.

## Limitations

- Input must be PowerPoint `.pptx`; legacy `.ppt` files are not supported.
- The converter extracts supported text and structured data, not slide appearance or layout.
- Images are not embedded or OCR-processed; only their alternative text is extracted.
- For shapes that cannot be processed, the converter may log a warning and continue.
- Files that cannot be opened are reported and skipped.

## Packaging

The included `PPTX_Converter.spec` configures PyInstaller to package `convert_gui.py` and the `python-pptx` data files. To build an executable from a Python environment with PyInstaller installed:

```powershell
py -m pip install pyinstaller
py -m PyInstaller PPTX_Converter.spec
```

The build output is written under `dist/`.
