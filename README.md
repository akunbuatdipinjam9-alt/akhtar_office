# Akhtar Office

Akhtar Office is a desktop productivity application that combines document editing, presentation creation, and source code editing in one workspace.

Built with Electron, HTML, CSS, and JavaScript.

## Features

### Documents

- Rich text editing
- Headings and text formatting
- Tables
- Images
- Page breaks
- Page size and margin controls
- Find and replace
- Focus mode
- DOCX export
- PDF export

### Presentations

- Multiple slides
- Text, shapes, and images
- Drag and resize elements
- Layer ordering
- Slide themes
- Grid and snap
- Speaker notes
- Slide transitions
- Presentation mode
- PPTX export

### Code Editor

- Line numbers
- Multi-tab editing
- Find and replace
- Go to line
- Word wrap
- Smart indentation
- JSON formatting
- Trailing whitespace cleanup
- Support for common source file formats

Supported formats include:

`HTML` `CSS` `JavaScript` `TypeScript` `JSON` `Python` `Markdown` `TXT`

## File Support

Akhtar Office can open several common file types.

| Format | Support |
| --- | --- |
| `.akoffice` | Native project format |
| `.docx` | Import and export |
| `.pptx` | Import and export |
| `.pdf` | Export |
| `.html` | Edit |
| `.css` | Edit |
| `.js` | Edit |
| `.ts` | Edit |
| `.json` | Edit |
| `.py` | Edit |
| `.md` | Edit |
| `.txt` | Edit |

DOCX and PPTX compatibility is still under development. Complex layouts, animations, and some formatting may not be preserved perfectly.

## Installation

### Requirements

- Windows 10 or Windows 11
- Node.js
- npm

Clone or download this repository, then install dependencies:

```bash
npm install
```

Start Akhtar Office:

```bash
npm start
```

Windows users can also run:

```text
RUN-WINDOWS.cmd
```

## Build for Windows

Run:

```text
BUILD-WINDOWS.cmd
```

Or manually:

```bash
npm run dist
```

The generated installer will be available inside:

```text
release/
```

## Windows File Associations

The Windows installer registers Akhtar Office as an available application for several supported file formats.

After installation, right-click a supported file:

```text
Open with
→ Choose another app
→ Akhtar Office
```

Windows controls the default application selection. Akhtar Office does not automatically replace existing default applications.

## Keyboard Shortcuts

| Shortcut | Action |
| --- | --- |
| `Ctrl + S` | Save |
| `Ctrl + O` | Open |
| `Ctrl + N` | New document |
| `Ctrl + W` | Close tab |
| `Ctrl + B` | Toggle sidebar |
| `Ctrl + Shift + P` | Command palette |
| `Ctrl + Shift + F` | Document focus mode |

## Project Structure

```text
akhtar-office/
├── index.html
├── styles.css
├── renderer.js
├── preload.js
├── main.js
├── package.json
├── RUN-WINDOWS.cmd
├── BUILD-WINDOWS.cmd
└── README.md
```

## Technology

- Electron
- HTML
- CSS
- JavaScript
- Node.js
- PptxGenJS
- docx
- Mammoth
- JSZip

## Design

Akhtar Office uses a minimal desktop interface focused on productivity.

The interface avoids unnecessary visual effects and keeps editing tools close to the active workspace.

The application does not include AI features.

## Current Limitations

Akhtar Office is still under active development.

Known limitations include:

- DOCX import is not fully compatible with every Microsoft Word layout.
- PPTX import currently has limited support for complex slide designs.
- Some PowerPoint animations are not supported.
- The code editor does not currently include a language server.
- Debugging and extension support are not available yet.

## Development Status

Current version:

```text
v11
```

Main development areas:

- Document compatibility
- Presentation compatibility
- Code editor improvements
- Workspace improvements
- Performance and stability
- Windows integration

## License

License information will be added before public release.

---

Akhtar Office

Documents. Presentations. Code.
