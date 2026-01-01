# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ProSchedules is a Revit add-in that provides tools for automating sheet management tasks (Duplication, Renaming). The add-in supports both Revit 2024 (.NET Framework 4.8) and Revit 2026 (.NET 8.0-windows) through multi-targeting.

## Build & Development Commands

### Building the Project
```bash
# Build for all target frameworks (net48 and net8.0-windows)
dotnet build ProSchedules.csproj

# Build for specific framework
dotnet build ProSchedules.csproj -f net48
dotnet build ProSchedules.csproj -f net8.0-windows

# Build the solution
dotnet build ProSchedules.sln
```

### Output Location
The project uses a custom output path: `C:\Users\mibil\OneDrive\Desktop\DevDlls\ProSchedules`
This is configured in the `<BaseOutputPath>` property in ProSchedules.csproj.

### Revit API References
- Revit 2024 (net48): References DLLs from `E:\Autodesk\Revit 2024\`
- Revit 2026 (net8.0-windows): References DLLs from `E:\Revit 2026\`

## Architecture Overview

### Command-Event-Service Pattern
The add-in follows a strict separation between UI commands, external events, and business logic:

1. **Commands** (`Commands/`): Entry points triggered by Revit ribbon buttons
   - `DuplicateSheetsCommand.cs`: Opens the Sheet Manager window (singleton pattern)
   - Uses Win32 interop to properly parent windows to Revit's main window

2. **External Events** (`ExternalEvents/`): Handlers that execute Revit API operations
   - `SheetDuplicationHandler.cs`: Duplicates sheets with configurable modes
   - `SheetEditHandler.cs`: Renames/updates sheets in a transaction
   - `SheetDeleteHandler.cs`: Deletes sheets in a transaction
   - External events are required because Revit API operations must run on Revit's main thread

3. **Services** (`Services/`): Business logic separated from UI
   - `RevitService.cs`: Core logic for retrieving sheets

### Sheet Duplication Modes

The sheet duplication feature supports three modes defined in `SheetDuplicationHandler.cs`:

- `EmptySheet`: Creates empty sheets with no views or detailing
- `WithSheetDetailing`: Includes detail lines, text notes, annotations, dimensions, tags
- `WithViews`: Includes detailing plus duplicates all placed views (model views, legends, schedules)

Additional options:
- `KeepLegends`: Places same legend views on duplicated sheet (not duplicated, reused)
- `KeepSchedules`: Places same schedule instances on duplicated sheet
- `CopyRevisions`: Copies revision information
- `CopyParameters`: Copies sheet and title block instance parameters

The handler implements a fallback strategy:
1. Attempts standard `ViewSheet.Duplicate()` with specified mode
2. Falls back to `ViewDuplicateOption.Duplicate` if WithDetailing fails
3. Manual sheet creation if both fail
4. Custom detailing copy via `CopySheetDetailing()` using `ElementTransformUtils.CopyElements()`

### UI Architecture

WPF windows with Material Design theming:

- **DuplicateSheetsWindow** (`UI/DuplicateSheetsWindow.xaml`): Main interface (Sheet Manager)
  - Searchable DataGrid of sheets
  - Duplication options
  - Bulk renaming (Find/Replace, Prefix/Suffix)
  - Deletion
  - Theme toggling

- **TitleBar** (`UI/TitleBar.xaml`): Reusable custom title bar component
  - Window dragging
  - Minimize/close buttons
  - Consistent styling across windows

Theme system uses dynamic resource dictionaries that can switch between light/dark modes at runtime.

### Dependency Management

The project uses Costura.Fody to embed dependencies into the output DLL, configured via `FodyWeavers.xml`. This ensures the add-in is distributed as a single DLL without dependency conflicts.

Key dependencies:
- MaterialDesignThemes/MaterialDesignColors: UI theming
- ricaun.Revit.UI: Revit ribbon/UI helpers
- netDxf, Newtonsoft.Json: Utility libraries

## Common Patterns

### Adding a New Command
1. Create command class in `Commands/` implementing `IExternalCommand`
2. Add `[Transaction(TransactionMode.Manual)]` attribute
3. Register in `App.cs` OnStartup via ribbon button
4. Create corresponding external event handler if Revit API operations needed
