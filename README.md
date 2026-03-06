# RibbonX Wizard

RibbonX Wizard is a Windows desktop tool for generating and applying Microsoft Word Ribbon customizations to macro-enabled templates (`.dotm`).

It helps users take an existing template with VBA macros and quickly add a custom Ribbon button that calls a selected macro, without manually editing Open XML package parts.

## Project Goal

The project’s goal is to simplify Word RibbonX customization by providing a guided workflow that:

1. Collects ribbon button settings (target tab, group label, button label, tooltip, icon, macro name).
2. Generates valid `customUI14.xml` (RibbonX) markup.
3. Generates matching VBA callback code (`onAction`) for the button.
4. Injects the XML into a copied output Office file (`.dotm`) safely.
5. Provides step-by-step instructions for completing setup in Word/VBA.

In short: **make Ribbon customization accessible to non-OpenXML specialists while preserving control over macro behavior.**

## Design Overview

The solution uses a **WPF + MVVM-style** architecture with clear separation of concerns:

- **UI layer** (`MainWindow.xaml`)
  - Presents a 4-step wizard-like experience: Configuration, Generated XML, VBA Code, and Instructions.
  - Uses data binding and command binding to interact with the ViewModel.

- **ViewModel layer** (`ViewModels/MainViewModel.cs`)
  - Owns all UI state and validation logic.
  - Orchestrates generation and save actions through commands:
    - Browse source file
    - Generate customization artifacts
    - Save output `.dotm`
    - Copy VBA and instructions to clipboard

- **Domain models** (`Models/`)
  - `RibbonConfig`: captures user input and derives stable IDs/callback names.
  - `GeneratedRibbon`: transport object containing generated XML, VBA code, and user instructions.

- **Generation service** (`Services/RibbonXmlGenerator.cs`)
  - Converts `RibbonConfig` into:
    - `customUI14.xml` ribbon content
    - VBA callback code that calls the user macro
    - human-readable setup instructions
  - Handles XML escaping and tab display-name mapping.

- **Document processing service** (`Services/DotmProcessor.cs`)
  - Validates input and output paths.
  - Opens Office package files and inserts/replaces RibbonX parts.
  - Uses `XmlPart.RibbonX14` for modern Office compatibility.

- **OpenXML package abstraction** (`OfficeRibbonXEditor.Common/`)
  - `OfficeDocument` and `OfficePart` encapsulate package operations:
    - create/read/update/remove custom UI parts
    - save modified Office files
    - inspect existing customization state

- **Static reference data** (`Data/BuiltInTabs.cs`)
  - Provides known Word tab `idMso` values exposed to the UI.

## End-to-End Flow

1. User selects a source `.dotm` file and enters button metadata.
2. `MainViewModel` builds a `RibbonConfig`.
3. `RibbonXmlGenerator` returns generated XML + VBA + instructions.
4. User reviews outputs in the wizard tabs.
5. On save, `DotmProcessor` writes RibbonX into a new output file.
6. User pastes generated VBA callback into the template and reopens Word.

## Technology Stack

- .NET 9 (`net9.0-windows`)
- WPF desktop UI
- C#
- CommunityToolkit.Mvvm dependency (project currently uses a lightweight internal command implementation in `MainViewModel`)

## Notes and Constraints

- Designed primarily for Word macro-enabled templates (`.dotm`) but extension checks also allow `.docm`, `.dotx`, and `.docx`.
- The target macro name must exist in the document’s VBA project.
- The app injects Ribbon XML; users still complete final VBA wiring inside Word’s VBA editor.
