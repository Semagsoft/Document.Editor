---
name: document.editor
description: Use when working with Document.Editor source code — document management, UI components, file conversion, and plugin architecture in this Qt/C++ project.
---

# Skills

This file defines the specialized skills available for the Document.Editor project.

## Document Management
Use this skill for tasks related to handling, loading, saving, and managing document tabs.

**Key Classes:**
- `DocumentManager`: Coordinates document lifecycle (create, open, save, close) and tab management.
- `DocumentEditor`: Handles content manipulation, file I/O for different formats, and formatting features.
- `DocumentTab`: Provides the UI container for the editor, including tab titles and context menus.

**Common Workflows:**
- **Opening/Creating a Document**: Triggered via `DocumentManager`, which creates a `DocumentTab` and instructs the `DocumentEditor` to load content.
- **Editing**: `DocumentEditor` updates the internal `QTextDocument`. Signals update the `DocumentTab` title and `DocumentManager` status bar.
- **Saving**: `DocumentManager` checks for unsaved changes in `DocumentEditor` and triggers a save or "Save As" dialog.
- **Closing**: `DocumentManager` prompts to save unsaved changes before removing the `DocumentTab`.

## UI Interaction
Use this skill for tasks related to manipulating the application's UI components and widgets.

**Key Classes:**
- `MainWindow`: The primary controller managing the main layout, toolbars, menu bar, and `QMdiArea`.
- `StatusBarManager`: Manages and updates the status bar with real-time statistics.
- `RulerWidget`, `ImageResizer`, `SymbolPicker`, `TableGridPicker`: Specialized widgets for ruler drawing, image resizing, and selection tools.

**Patterns:**
- **Mediator Pattern**: `MainWindow` delegates document complexity to `DocumentManager`.
- **Dynamic UI State**: Uses `updateEditorActions()` to toggle UI elements based on the active document's state.
- **Signal/Slot Communication**: Heavily uses Qt signals to decouple components (e.g., `SymbolPicker` emitting a selection signal).

## File Conversion
Use this skill for tasks involving file format detection and conversion between different formats.

**Supported Formats:**
- **XAML/DEXML**: Native XML format supporting complex document structures.
- **HTML**: Native support via Qt.
- **RTF**: Basic support (saved as HTML-wrapped content).
- **Plain Text**: Fallback format for simple text.

**Workflow:**
1. **Trigger**: User action (e.g., "Save As") provides a file path.
2. **Detection**: `DocumentEditor` identifies the target format by file extension.
3. **Conversion**:
    - `.xaml`: Calls `saveToXaml()` for XML serialization.
    - `.html`: Calls `toHtml()` for HTML export.
    - `.txt`: Calls `toPlainText()` for plain text.
4. **Completion**: `DocumentEditor` updates its state, and `DocumentManager` emits `documentSaved`.

## Plugin Management
Use this skill for tasks involving the loading, registration, and interaction with plugins.

**Architecture:**
- **`IPlugin` Interface**: All plugins must implement `name()`, `version()`, `description()`, `initialize()`, and `shutdown()`.
- **`PluginManager`**: Responsible for scanning the `/plugins` directory, loading shared libraries via `QPluginLoader`, and managing lifecycle calls.

**Interaction Patterns:**
- **Discovery**: Query `PluginManager` for plugin names and counts to dynamically build UI elements.
- **Access**: Retrieve `PluginInfo` to call plugin-specific methods via the `IPlugin` instance.
- **Extension**: Plugins are loaded dynamically from the filesystem, allowing for modular extensions without modifying core source code.
