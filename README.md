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
cmake -B build -DCMAKE_BUILD_TYPE=Release
cmake --build build
```

## Running the tests

```bash
QT_QPA_PLATFORM=offscreen ctest --test-dir build --output-on-failure
```

The suite needs no display; every test runs headlessly using Qt's offscreen
platform plugin.

## Installing / packaging

```bash
cmake --install build --prefix /usr/local   # binary, .desktop file, icon, docs
cpack --config build/CPackConfig.cmake      # DEB (Linux), TGZ; NSIS (Windows), DMG (macOS)
```

## Plugins

An example plugin lives in `plugins/example`. Plugins declare
`"apiVersion": 1` in their JSON metadata and are loaded only when enabled
in Options → Plugins. Plugins compiled against a different API version are
refused before any plugin code runs.

Note: the FTP import/export feature transfers in plaintext; the app warns
before using it.

## License

GPL v2 License — see LICENSE.
