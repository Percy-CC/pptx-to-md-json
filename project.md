# Project Overview

## Name

PPTX to AI

## Purpose

PPTX to AI is a local PowerPoint text-extraction tool. It converts `.pptx` presentations into Markdown and JSON corpora that are easier to search, review, and provide to AI tools. Conversion happens on the user's machine; the application does not upload presentations to an external service.

## Users and use cases

- Educators and content teams who want to review or reuse presentation text.
- AI users who need slide content in text-based formats for analysis or retrieval.
- Developers who want a structured, per-slide representation of PowerPoint content.

## Current functionality

- Processes `.pptx` files from an `input` directory or files selected in the desktop app.
- Preserves slide order and records slide number, title, hidden-slide status, and speaker notes.
- Extracts text and paragraph levels, tables, chart titles and series data, SmartArt text, image alternative text, and embedded-object labels.
- Represents common text formatting in Markdown, including bold, italic, underline, strikethrough, links, superscript, subscript, highlights, and non-default font colors.
- Writes a combined Markdown document and a JSON array to an `output` directory.
- Includes a Tkinter desktop interface for choosing presentations and an output folder, monitoring conversion logs, and opening the output folder.
- Includes a PyInstaller specification for packaging the desktop app as a Windows executable.

## Out of scope

- Rendering slides or reproducing their visual layout.
- Extracting text from image pixels with OCR.
- Embedding image files in the generated corpus; only available image alternative text is included.
- Editing or modifying source presentations.
- Sending files to an AI provider or making AI-generated summaries.

## Components

| File | Responsibility |
| --- | --- |
| `convert_pptx.py` | Reads presentations and produces Markdown and JSON output. |
| `convert_gui.py` | Provides the desktop interface and invokes the converter for selected files. |
| `PPTX_Converter.spec` | Configures PyInstaller packaging for the desktop app. |
| `input/` | Default directory for presentations when running the converter directly. |
| `output/` | Default destination for the generated corpus. |

## Data flow

1. The user places `.pptx` files in `input/` or chooses files in the desktop app.
2. The converter opens each presentation, walks through its slides and shapes, and extracts supported text and structured content.
3. The converter writes `all_slides.md` and `all_slides.json` to the selected output directory.

The GUI copies its selected presentations to a temporary working directory before invoking the converter. When two selected files have the same name, it adds a numeric suffix to avoid collisions.

## Output contract

`all_slides.md` is organized by source file and slide. It includes slide titles, hidden-slide markers, extracted content, and speaker notes where present.

`all_slides.json` contains one object per slide, with these fields:

| Field | Description |
| --- | --- |
| `source_file` | Original presentation filename. |
| `slide_number` | One-based slide number. |
| `hidden` | Whether the slide is marked hidden. |
| `title` | Slide title, or an empty string when none is available. |
| `blocks` | Ordered content blocks such as text, tables, charts, SmartArt, images, and embedded objects. |
| `speaker_notes` | Extracted speaker notes, or an empty string. |

## Operational notes

- The direct converter reads only `.pptx` files immediately inside `input/`; it does not recurse into subdirectories.
- The desktop app can add individual `.pptx` files or the `.pptx` files immediately inside a chosen folder.
- Output filenames are fixed. A new conversion replaces existing `all_slides.md` and `all_slides.json` files in the chosen output folder.
- A presentation that cannot be opened is reported and skipped; unsupported or problematic individual shapes may also be skipped with a warning.
- The implementation is Python and uses `python-pptx`, `lxml`, and the standard-library Tkinter GUI toolkit.

## Future direction

Keep conversion local, preserve a predictable per-slide JSON structure, and improve extraction coverage without attempting to reproduce slide visuals. Changes to output fields or content-block types should be reflected in this document and in the README.
