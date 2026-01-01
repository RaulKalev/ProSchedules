# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PlaceViews is a Revit add-in that provides tools for automating view and sheet management tasks. The add-in supports both Revit 2024 (.NET Framework 4.8) and Revit 2026 (.NET 8.0-windows) through multi-targeting.

## Build & Development Commands

### Building the Project
```bash
# Build for all target frameworks (net48 and net8.0-windows)
dotnet build PlaceViews.csproj

# Build for specific framework
dotnet build PlaceViews.csproj -f net48
dotnet build PlaceViews.csproj -f net8.0-windows

# Build the solution
dotnet build PlaceViews.sln
```

### Output Location
The project uses a custom output path: `C:\Users\mibil\OneDrive\Desktop\DevDlls\PlaceViews`
This is configured in the `<BaseOutputPath>` property in PlaceViews.csproj.

### Revit API References
- Revit 2024 (net48): References DLLs from `E:\Autodesk\Revit 2024\`
- Revit 2026 (net8.0-windows): References DLLs from `E:\Revit 2026\`

## Architecture Overview

### Command-Event-Service Pattern
The add-in follows a strict separation between UI commands, external events, and business logic:

1. **Commands** (`Commands/`): Entry points triggered by Revit ribbon buttons
   - `MainCommand.cs`: Opens the Place Views window (singleton pattern)
   - `DuplicateSheetsCommand.cs`: Opens the Duplicate Sheets window (singleton pattern)
   - Both implement singleton pattern to prevent duplicate windows
   - Both use Win32 interop to properly parent windows to Revit's main window

2. **External Events** (`ExternalEvents/`): Handlers that execute Revit API operations
   - `ViewPlacementHandler.cs`: Places selected views onto selected sheets
   - `SheetDuplicationHandler.cs`: Duplicates sheets with configurable modes
   - External events are required because Revit API operations must run on Revit's main thread
   - WPF UI runs asynchronously, so commands are queued via ExternalEvent.Raise()

3. **Services** (`Services/`): Business logic separated from UI and Revit threading
   - `RevitService.cs`: Core logic for view/sheet operations and fuzzy matching algorithm

### View Placement Logic

The view-to-sheet matching uses a sophisticated fuzzy matching algorithm in `RevitService.cs`:

1. First attempts exact case-insensitive match
2. Extracts key parts from sheet names (splits by separators: space, dash, underscore, etc.)
3. Scores each view based on:
   - All sheet parts must be present in view name (base requirement)
   - High bonus (+100) for exact suffix match (critical for "Heli B01-c" vs "Heli B01-d")
   - Medium bonus (+50) for suffix appearing near end
   - Additional points (+5 each) for each matching part
4. Returns highest-scoring view

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

- **MainWindow** (`UI/MainWindow.xaml`): Two-panel interface with searchable DataGrids
  - Left panel: Available views with search and multi-select
  - Right panel: Sheets with search and multi-select
  - Custom title bar with theme toggle
  - Manual window resizing via edge/corner mouse handlers
  - Popup overlay for displaying results

- **DuplicateSheetsWindow**: Sheet duplication interface with mode selection

- **TitleBar** (`UI/TitleBar.xaml`): Reusable custom title bar component
  - Window dragging
  - Minimize/close buttons
  - Consistent styling across windows

Theme system uses dynamic resource dictionaries that can switch between light/dark modes at runtime.

### Data Models

Simple POCO models in `Models/`:

- `ViewItem`: Wraps Revit View with name, type, element ID, and selection state
- `SheetItem`: Wraps Revit ViewSheet with name, number, element ID, and selection state

Both implement `INotifyPropertyChanged` for WPF data binding, particularly for the `IsSelected` property used in DataGrid checkboxes.

### Dependency Management

The project uses Costura.Fody to embed dependencies into the output DLL, configured via `FodyWeavers.xml`. This ensures the add-in is distributed as a single DLL without dependency conflicts.

Key dependencies:
- MaterialDesignThemes/MaterialDesignColors: UI theming
- ricaun.Revit.UI: Revit ribbon/UI helpers
- netDxf, Newtonsoft.Json: Utility libraries

## Important Implementation Notes

### Thread Safety
- Never call Revit API directly from WPF event handlers
- Always use IExternalEventHandler pattern for Revit operations
- UI updates after external events must be marshaled back to UI thread

### Window Management
- Commands maintain static window references to implement singleton pattern
- Win32 interop (`SetForegroundWindow`, `ShowWindow`) ensures proper window activation
- Windows are parented to Revit's main process window handle

### Transaction Handling
- All Revit document modifications must occur within a Transaction
- External event handlers create and manage their own transactions
- Rollback on exception is automatic when transaction is not committed

### Platform Targeting
- Use `#if NET48` or `#if NET8_0_OR_GREATER` for version-specific code
- Most Revit API patterns work identically across both frameworks
- Be aware of .NET API differences (e.g., newer C# features in .NET 8)

### Element ID Usage
- Prior to Revit 2024, ElementId was a reference type
- Revit 2024+ ElementId has `.Value` property (long)
- Code uses `.Value` for category ID comparisons in `SheetDuplicationHandler.cs:348`

## Testing in Revit

The add-in registers under the "RK Tools" tab in Revit's ribbon (configured in `App.cs`). To test:

1. Build the project for appropriate target framework
2. Output DLL will be in the custom output path
3. Ensure Revit add-in manifest (.addin file) points to correct DLL location
4. Launch Revit and look for "RK Tools" tab
5. Click "Place Views" or "Duplicate Sheets" buttons

## Common Patterns

### Adding a New Command
1. Create command class in `Commands/` implementing `IExternalCommand`
2. Add `[Transaction(TransactionMode.Manual)]` attribute
3. Register in `App.cs` OnStartup via ribbon button
4. Create corresponding external event handler if Revit API operations needed

### Adding a New Window
1. Create XAML + code-behind in `UI/`
2. Reference Material Design themes in Window.Resources
3. Use TitleBar control for consistent chrome
4. Implement singleton pattern in command if appropriate
5. Set window owner to Revit process using WindowInteropHelper

### Modifying Fuzzy Matching
The matching algorithm is in `RevitService.cs`:
- `FindBestMatch()`: Main entry point
- `CalculateMatchScore()`: Scoring logic
- `ExtractKeyParts()`: Name parsing
Adjust scoring weights to change matching behavior.
