# Document.Editor

A cross-platform rich text document editor built with Qt6 and C++17.

## Features

- Rich text editing with XAML, HTML, RTF, and TXT support
- Multi-tab/MDI document interface
- Text-to-speech
- Spell checking
- FTP import/export
- Plugin architecture
- Multiple themes (Office 2010, Office 2013, Windows 8, Silver, Black)
- Page layout and print support
- Shape and chart insertion
- Word count, find/replace, and other editing tools

## Dependencies

- CMake 3.20+
- C++17 compiler
- Qt6 (Core, Gui, Widgets, PrintSupport, Network, TextToSpeech, Xml, Charts, Concurrent)

## Building

```bash
mkdir build && cd build
cmake ..
cmake --build .
```

## License

GPL v2 License — see LICENSE.
