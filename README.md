# Papyr

Papyr is a lightweight, read-only DOCX and XLSX viewer built with Rust and Tauri. The goal is simple: open Office documents quickly, render useful content, and avoid the overhead of launching a full editor when you only need to read.

## Installation

This repository currently documents installation from source. You will need:

- Rust and Cargo
- Tauri prerequisites for your platform
- on Windows, Microsoft Edge WebView2

1. Install the Tauri CLI:

```bash
cargo install tauri-cli
```

2. Build the Papyr desktop app from the repository root:

```bash
cargo tauri build
```

You can also run the same command from `src-tauri/`:

```bash
cd src-tauri
cargo tauri build
```

If you want to build the Rust desktop crate directly without Tauri bundling, use plain Cargo instead:

```bash
cargo build --manifest-path src-tauri/Cargo.toml
```

## How to Use

1. Launch Papyr.
2. Open a `.docx` or `.xlsx` file from the welcome screen or the toolbar.
3. You can also drag and drop a `.docx` or `.xlsx` file onto the window, or reopen a file from the recent files list.
4. Use the toolbar or keyboard shortcuts to search the document, show comments for DOCX files, and switch themes.

Useful shortcuts:

- `Ctrl+O` open a file
- `Ctrl+F` find in the current document
- `Ctrl++` / `Ctrl+-` zoom the document in or out
- `Ctrl+Mouse Wheel` zoom the document in or out
- `Ctrl+]` show or hide comments
- `Ctrl+D` toggle light and dark theme
- `Ctrl+Q` quit Papyr

## What Papyr Supports

Papyr is a read-only viewer and currently supports:

- `.docx` documents
- paragraphs and styled text runs
- tables
- explicit page breaks
- comments
- headers and footers
- footnotes
- embedded images
- `.xlsx` workbooks
- sheet tabs
- fast spreadsheet grid previews
- shared strings, inline strings, numbers, booleans, dates, formulas with cached values, and basic number/date/text formats
- recent files

## Useful Notes

- Papyr opens `.docx` and `.xlsx` files for reading only. It does not edit or save documents.
- XLSX support is a fast preview path. It does not evaluate formulas or render charts, pivot tables, macros, embedded objects, or full Excel-compatible layout.
- Recent files are stored locally by the desktop app so you can reopen them quickly.
- For architecture, development notes, and current technical limitations, see [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).
- The product specification lives in [docs/SPEC.md](docs/SPEC.md).

## License

No license file is currently included in this repository.
