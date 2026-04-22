# AnyView

macOS native document viewer. View anything, locally, fast, faithful.

No Office, no Electron, no cloud. Just drop a file and see it.

| Light | Dark |
|-------|------|
| ![Code - Light](screenshots/code-swift.png) | ![Code - Dark](screenshots/code-dark.png) |
| ![Markdown](screenshots/markdown.png) | |

## Supported Formats

| Category | Formats | Engine |
|----------|---------|--------|
| PDF | pdf | PDFKit |
| Word | docx, docmod, doct | docmod CLI |
| Presentations | pptx, ppt, key | Quick Look |
| Spreadsheets | xlsx, xls, numbers | Quick Look |
| Pages | pages | Quick Look |
| Images | png, jpg, gif, webp, tiff, bmp, ico, heic, svg | NSImageView |
| Markdown | md, markdown | highlight.js + mermaid |
| HTML | html, htm | WKWebView |
| Code | 60+ languages | highlight.js |
| Data/Config | json, yaml, toml, xml, csv, plist, ini... | highlight.js |

## Features

- Drag & drop to open
- Multi-tab browsing
- Zoom 50%--300%
- Light / Dark theme
- Auto-reload on file change
- Mermaid diagrams in Markdown
- HTML preview / source toggle
- Syntax highlighting for code

## Build

```bash
# requires macOS 13+, Swift 5.9+
swift build

# build .app bundle
./build-app.sh
```

## Architecture

Pluggable renderer protocol -- adding a new format is one file:

```
ViewerRenderer (protocol)
  |-- PDFRenderer        (PDFKit)
  |-- ImageRenderer      (NSImageView)
  |-- QuickLookRenderer  (QLPreviewView)
  |-- WebRenderer        (WKWebView -- docx/md/html/code)
```

`ViewerWindowController` handles window, toolbar, file watching, zoom. Renderers handle rendering. ~213 lines of controller, zero `if isFormatX` branches.

## Requirements

- macOS 13+
- [docmod](https://github.com/cove-apps/docmod) CLI (for docx/docmod/doct rendering, optional)
