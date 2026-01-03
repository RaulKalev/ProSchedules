# ProSchedules

ProSchedules is a powerful Revit add-in designed to streamline sheet management tasks. It automates sheet duplication, renaming, and organization, significantly reducing manual effort for Revit professionals.

The add-in supports both **Revit 2024** (.NET Framework 4.8) and **Revit 2026** (.NET 8.0-windows) through multi-targeting.

## Features

### 📋 Advanced Sheet Duplication
- **Multiple Modes**:
  - **Empty Sheet**: Creates new empty sheets.
  - **With Detailing**: Copies detail lines, text, annotations, tags, and dimensions.
  - **With Views**: Duplicates sheets along with all placed views (model views, legends, schedules).
- **Control Options**:
  - Keep Legends & Schedules (reuse existing instances).
  - Copy Revisions & Parameters (sheet/titleblock data).
  - Specify number of copies.

### 🔄 Bulk Renaming
- **Find & Replace**: Batch rename sheets using find/replace logic.
- **Prefix/Suffix**: Add prefixes or suffixes to sheet numbers and names.
- **Auto-Numbering**: Intelligent renumbering of duplicated sheets.

### 📊 Schedule Management & Sorting
- **Advanced Sorting**:
  - Sort sheets/schedules by any available column parameter.
  - Multi-level sorting (add multiple criteria).
  - Custom Sorting Window with persistent settings per schedule.
  - "Apply" without closing to test sort logic.
  - Drag-resizing and movable windows for better UX.
- **Grouping**: Option to itemize every instance or group by sorting keys.
- **UI Visuals**:
  - Yellow highlight for modified rows.
  - Clear selection feedback.

### 🖱️ Excel-Like Interaction
- **Smart Selection**:
  - Drag-select cells with a cleaner, perimeter-only selection border (no cluttered full-grid highlight).
  - Handles large selections smoothly with virtualization support.
- **Auto-Fill & Copy**:
  - **Sequential Fill**: Drag the fill handle to automatically increment numbers (e.g., "Sheet 1" → "Sheet 2").
  - **Copy Mode**: Hold `Ctrl` while dragging to perform an exact copy.
  - **Visual Feedback**: Dynamic crosshair cursor and "Plus" indicator for copy mode.

### 🔍 Highlight in Model (New!)
- **Smart Selection**:
  - Tick checkboxes to select items (supports both **Sheet Lists** and **Schedule Items**).
  - **Group Selection**: Selecting a grouped row automatically selects **all** items in that group.
- **View Management**:
  - **Smart View Finding**: Prioritizes finding and opening a **Floor Plan** first.
  - **View Toggling**: Clicking again toggles between **Floor Plan** and **3D View**.
  - **Focus**: Selects and zooms to fit elements (keeping context visible).
- **Checkbox Autofill**:
  - Drag the fill handle while holding `Ctrl` to fast-check/uncheck multiple rows.

### 🔔 Smart Update System
- **Update Notification**: Automatically detects new versions and displays a rich "What's New" changelog popup on startup.
- **Project-Aware Settings**:
  - Persists "Last Selected Schedule" and "Sort Criteria" specifically for each Revit project.
  - Automatically restores your preferred workflow when you switch between projects.
- **Robust Error Handling**:
  - Automatically reverts invalid parameter edits (e.g., values rejected by Revit) to keep the DataGrid in sync with the model.

### 🛡️ Security
- **Code Signing**: All Release builds are automatically signed with a self-signed certificate (`RKToolsCert.pfx`), establishing trust and preventing "Publisher Unknown" warnings in Revit.

### 🎨 Modern UI/UX
- **Material Design**: Sleek, dark-themed interface using Material Design for XAML.
- **Custom Window Chrome**: Borderless windows with custom title bars, resizing, and docking.
- **Search**: Real-time filtering of sheets.
- **Theme Toggling**: Switch between Light and Dark modes (WIP).

## Installation

1. Clone the repository.
2. Build the project using the commands below.
3. The DLLs are output to `C:\Users\mibil\OneDrive\Desktop\DevDlls\ProSchedules` (configurable in `.csproj`).
4. Load the add-in in Revit.

## Development

### Prerequisites
- Visual Studio 2022 or newer (for .NET 8.0 support).
- Revit 2024 and/or Revit 2026 SDK/DLLs installed.

### Build Commands
```bash
# Build for all targets
dotnet build ProSchedules.csproj

# Build for specific version
dotnet build ProSchedules.csproj -f net48       # Revit 2024
dotnet build ProSchedules.csproj -f net8.0-windows # Revit 2026
```

## Architecture
- **Commands**: Entry points (`DuplicateSheetsCommand`) handling UI invocation.
- **External Events**: `IExternalEventHandler` implementations (`SheetDuplicationHandler`, `SheetEditHandler`) for safe Revit API interaction.
- **Services**: `RevitService` for data retrieval and logic.
- **UI**: WPF MVVM-like pattern with `DuplicateSheetsWindow` and `SortingWindow`.
- **Resources**: Global styling in `ElementStyles.xaml` for consistency.

## Dependencies
- **MaterialDesignInXaml**: UI Styling.
- **ricaun.Revit.UI**: Ribbon integration.
- **Costura.Fody**: DLL embedding for single-file distribution.

## Known Issues / TODO

### Parameters Window
- **Built-in Type Parameter Values**: Some built-in type parameters (e.g., "Model", "Description") do not currently display their values in the DataGrid, although they show correctly in Revit's native schedule editor. Instance parameters, shared parameters, and project parameters work correctly.
  - Status: Under investigation
  - Workaround: Use Revit's native schedule editor for viewing these specific parameter values

---

