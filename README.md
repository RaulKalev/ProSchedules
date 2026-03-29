# ProSchedules

ProSchedules is a powerful Revit add-in designed to streamline schedule management tasks. It provides an advanced schedule viewer with sorting, filtering, grouping, inline editing, and Excel-like interactions — significantly reducing manual effort for Revit professionals.

The add-in supports both **Revit 2024** (.NET Framework 4.8) and **Revit 2026** (.NET 8.0-windows) through multi-targeting.

## Features

### 📊 Schedule Viewer
- Browse all schedules in the active Revit project from a single window.
- **Itemize / Group** toggle: view every instance individually or collapse rows by sort key.
- **`<Varies>`** display: grouped rows with differing values across instances clearly show `<Varies>` for those columns.
- **Real-time search**: filter visible rows across all schedule columns instantly.

### 🔃 Advanced Sorting
- Sort by any available schedule column parameter.
- Multi-level sorting — add multiple criteria and drag to reorder their priority.
- **Blank line rows**: insert Revit-style separator rows between sort groups.
- **Footer rows**: append a footer after each sort group with configurable display options:
  - *Title, count, and totals* / *Title and totals* / *Count and totals* / *Totals only*
- Settings persist per schedule via **Revit Extensible Storage** (worksharing-safe).
- "Apply" without closing to preview results live.

### 🔎 Advanced Filtering
- Filter rows by any parameter column using 14 Revit-style conditions:
  `equals`, `does not equal`, `contains`, `does not contain`, `begins with`, `ends with`, `greater than`, `less than`, `has a value`, `has no value`, and more.
- Value field is an editable ComboBox — pick from existing values or type a custom value.
- Multi-rule filtering (all rules apply as AND conditions).
- Settings persist per schedule via **Revit Extensible Storage**.
- Rules are only committed when clicking **Apply** — Cancel always restores the previous state.

### 🖱️ Excel-Like Interaction
- **Inline editing**: edit parameter values directly in the grid; invalid edits are automatically reverted.
- **Smart selection**: drag-select cells with a perimeter-only selection border.
- **Auto-fill & copy**:
  - **Sequential fill**: drag the fill handle to automatically increment numbers (e.g., "Item 1" → "Item 2").
  - **Copy mode**: hold `Ctrl` while dragging for an exact copy.
  - **Visual feedback**: dynamic crosshair cursor and "Plus" indicator for copy mode.
- **Checkbox fill**: drag-fill checkboxes to check/uncheck multiple rows rapidly.

### 🔍 Highlight in Model
- Tick checkboxes to select schedule items in Revit.
- **Group selection**: selecting a grouped row automatically selects all items in that group.
- **Smart view finding**: prioritises opening a Floor Plan; toggling switches between Floor Plan and 3D View.
- **Focus**: selects and zooms to fit elements in the active view.

### 💾 Persistent Settings
- **Sort & filter settings** stored in Revit Extensible Storage per user per project — safe with worksharing/Revit Server.
- **Last selected schedule** remembered per project (local file).
- Settings are restored automatically when switching between projects or reopening Revit.

### 🎨 Modern UI/UX
- **Material Design**: sleek interface with switchable Light and Dark themes.
- **Theme-aware ribbon icon**: the ribbon button automatically matches Revit's current UI theme (dark/light).
- **Custom window chrome**: borderless windows with custom title bars, drag-to-move, and resize.
- **Sorting window**: drag-and-drop cards to reorder sort criteria.

### 🛡️ Security
- **Code signing**: all Release builds are automatically signed with a self-signed certificate (`RKToolsCert.pfx`), preventing "Publisher Unknown" warnings in Revit.

---

## Installation

1. Clone the repository.
2. Build the project (see commands below).
3. DLLs are output to `C:\Users\mibil\OneDrive\Desktop\DevDlls\ProSchedules` (configurable in `.csproj`).
4. Copy the appropriate `.addin` file and DLL to your Revit add-ins folder or load via the Revit Add-in Manager.

---

## Development

### Prerequisites
- Visual Studio 2022 or newer (for .NET 8.0 support).
- Revit 2024 and/or Revit 2026 installed (for API DLLs).

### Build Commands
```bash
# Build for all targets
dotnet build ProSchedules.csproj

# Build for a specific Revit version
dotnet build ProSchedules.csproj -f net48            # Revit 2024
dotnet build ProSchedules.csproj -f net8.0-windows   # Revit 2026
```

---

## Architecture
- **Commands** — Entry points implementing `IExternalCommand`, invoked from the Revit ribbon.
- **External Events** — `IExternalEventHandler` implementations for safe Revit API calls on the main thread (parameter updates, highlight, settings save).
- **Services** — `RevitService` and `ExtensibleStorageService` for data retrieval and persistent storage.
- **UI** — WPF windows: `DuplicateSheetsWindow` (main viewer), `SortingWindow`, `FilterWindow`.
- **Themes** — Dynamic resource dictionaries (`DarkTheme.xaml`, `LightTheme.xaml`) swapped at runtime.

## Dependencies
- **MaterialDesignInXaml** — UI styling.
- **ricaun.Revit.UI** — Ribbon integration helpers.
- **Costura.Fody** — Embeds all dependencies into a single DLL for clean distribution.
- **Newtonsoft.Json** — Sort/filter settings serialisation.

---

## Known Issues

- **Built-in type parameter values**: Some built-in type parameters (e.g., *Model*, *Description*) do not display their values in the grid, although they show correctly in Revit's native schedule editor. Instance parameters, shared parameters, and project parameters work correctly.
  - *Status*: Under investigation.
  - *Workaround*: Use Revit's native schedule editor for these specific parameters.


