# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

# ACTIVE FEATURE IMPLEMENTATION INSTRUCTIONS
## Retag / Normalize Existing Tags Workflow (SmartTags)

You are Claude Code working in an existing Revit add-in project.

GOAL  
Implement a “Retag / Update existing tags” workflow that allows SmartTags to act as a maintenance tool.

Two modes must be supported:
1) **Fully Automatic** – apply all adjustments silently and report results.
2) **User Confirmation** – step through each proposed adjustment, focus the view, and let the user Accept / Reject / Cancel.  
   - Changes are applied immediately.
   - If rejected, that specific change must be reverted to its previous state.
   - Cancel stops the process and reverts the current candidate only.

UI REQUIREMENTS
- Add a new **card in the right column, at the bottom** of the main placement window.
- Card contents:
  - Two **mutually exclusive** checkboxes:
    - “Fully automatic”
    - “User confirmation”
  - Buttons:
    - “Retag Selected”
    - “Normalize View”
- Default mode: Fully automatic (unless preferences say otherwise).
- Persist the selected mode in user preferences.

FUNCTIONAL REQUIREMENTS

A) Retag Selected  
- Use current selection (elements) in active view.
- Find all **SmartTags-managed tags** that reference those elements.
- For each tag:
  - Compute a new optimal placement using current SmartTags settings:
    - Direction
    - Leader settings
    - Rotation/orientation policy
    - Collision detection
  - If placement is unchanged within tolerance → skip.
- Execution modes:
  - Fully automatic:
    - Apply all changes in a batch transaction.
    - Show summary: adjusted / unchanged / skipped / failed.
  - User confirmation:
    - For each proposed change:
      - Focus the view on the tag location.
      - Prompt: Apply this adjustment? [Accept] [Reject] [Cancel]
      - Accept → keep change
      - Reject → revert that tag to previous state
      - Cancel → stop; revert current candidate only

B) Normalize View  
- Collect **all SmartTags-managed tags in the active view**.
- Run the same logic as Retag Selected.
- Same two execution modes.

SMARTTAGS MARKER (MANDATORY)
- Use **Extensible Storage on the tag element** to mark SmartTags-managed tags.
- Do NOT modify tag families or require shared parameters.
- Stored metadata must include:
  - Schema GUID
  - Plugin name + version
  - Creation timestamp
  - Referenced element id (if available)
  - managed = true
- New tags created by SmartTags must always write this marker.
- Existing tags without marker are treated as unmanaged and skipped.

ARCHITECTURE CONSTRAINTS
- Follow **Command → ExternalEvent → Service** separation strictly.
- No Revit API calls from WPF/UI thread.
- All model changes via ExternalEvent handlers.
- Confirmation flow must be implemented using **ExternalEvent-based state machine**, not blocking UI.
- Dialogs must be parented to Revit main window using existing Win32 helper.
- Respect existing Material Design theming.

IMPLEMENTATION ORDER (MANDATORY)

1) Data model
- Create TagAdjustmentProposal:
  - TagId
  - ReferencedElementId
  - OldStateSnapshot
  - NewStateProposal
  - Notes / reason
- Create TagStateSnapshot capturing everything needed to revert:
  - TagHeadPosition
  - Leader state you modify
  - Rotation/orientation you modify
  - Any parameters you touch

2) Extensible Storage
- Implement SmartTagMarkerStorage:
  - EnsureSchema()
  - SetManagedTag()
  - IsManagedTag()
  - TryGetMetadata()
- Integrate into existing tag placement handlers.

3) Tag discovery
- Implement service:
  - FindTagsReferencingElements(doc, view, elementIds)
- Filter:
  - Active view only
  - Managed tags only
- Handle Revit 2024 / 2026 API differences via #if.

4) Adjustment computation
- Compute proposed placements WITHOUT modifying model.
- Reuse collision engine.
- Return unchanged/skipped when within tolerance.

5) Automatic application
- ExternalEvent handler:
  - RetagApplyHandler
- One transaction per operation:
  - “SmartTags: Retag”
  - “SmartTags: Normalize View”
- Apply proposals, collect results, report to UI.

6) Confirmation workflow
- Implement RetagConfirmationController (UI-level coordinator).
- Use:
  - ApplySingleProposalHandler
  - RevertSingleProposalHandler
- Flow:
  - Apply proposal → focus view → ask user
  - Reject → revert snapshot
  - Cancel → revert current and stop
- Never leave model in partial/unknown state.

7) UI + Preferences
- Add card to bottom-right of window.
- Enforce checkbox exclusivity in code.
- Persist mode selection in preferences.

8) Safety & Edge Cases
- Handle deleted tags/elements gracefully.
- Skip invalid references.
- Use tolerance (~0.5 mm) for unchanged detection.
- Avoid infinite loops.
- Cache obstacles per run where possible.

DELIVERABLES
- UI updates
- New ExternalEvent handlers
- Services for discovery and computation
- Extensible Storage implementation
- Preference updates
- README section explaining Retag / Normalize

CODING RULES
- Small, readable methods
- No removal of existing functionality
- Multi-targeting must remain intact
- Explicit System.Windows.Visibility usage where applicable
- No emojis in code or comments

---

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
