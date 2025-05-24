# RSAP_FootingLocker

RSAP_FootingLocker is a VBA tool designed for modeling the geometry of footings in Autodesk Robot Structural Analysis Professional. The tool provides Excel macros and supporting VBA modules to streamline the creation and manipulation of footing geometries for structural engineering analysis.

## Features

- Automates the generation of various types of footings (column, wall, slab) for import into Robot Structural Analysis.
- User-friendly interface via Excel macros.
- Modular VBA codebase for easy customization and extension.

## Getting Started

### Prerequisites

- Microsoft Excel (with macro support enabled)
- Autodesk Robot Structural Analysis Professional (for integration)
- The `.xlsm` workbook (`FootingLocker-v1.40.xlsm`) and supporting VBA modules/classes

### Installation

1. Clone or download this repository.
2. Open `FootingLocker-v1.40.xlsm` in Excel.
3. Ensure macros are enabled.
4. Optionally, review or modify the VBA modules (`.bas`, `.cls`, `.doccls` files) via the Visual Basic for Applications editor.

### Usage

1. Open the macro-enabled Excel workbook.
2. Use the provided user interface to define and generate footing geometries.
3. Export or integrate the generated data with Robot Structural Analysis.

## Repository Structure

- `FootingLocker-v1.40.xlsm` – Main Excel workbook with macros and UI.
- `.bas`, `.cls`, `.doccls` files – VBA modules and class definitions for geometry modeling.
- Example files:
    - `CColumnFootingDTO.cls`
    - `CWallFootingDTO.cls`
    - `MConstants.bas`
    - `MRSAPUtilities.bas`
    - `Sheet1.doccls` (and others)

For a full list of files, see the [repository contents](https://github.com/teixeiranh/RSAP_FootingLocker/).

## Contributing

Contributions, issues, and feature requests are welcome! Please open an issue to discuss potential changes or enhancements.

## Acknowledgments

- Developed for streamlining geometry modeling in Robot Structural Analysis
- VBA and Excel macros for structural engineering applications

---

For more details or to browse code modules, visit the [repository on GitHub](https://github.com/teixeiranh/RSAP_FootingLocker/).
