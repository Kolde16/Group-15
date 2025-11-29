# 🌡️ IFC Thermal Quantifier

A Python tool that automates the extraction of building envelope data from IFC (Industry Foundation Classes) models. It parses Windows, Walls, Slabs, and Roofs to generate a comprehensive **Thermal Report** in Excel format, calculating weighted average U-values and material quantities.

## 🚀 Features

*   **Automated Extraction:** Scans IFC models for Walls, Windows, Slabs, and Roofs.
*   **Smart Data Cleaning:** Converts messy text inputs (e.g., "150mm", "0.3 W/m²K") into usable float numbers.
*   **Window Logic:** Calculates composite U-values by separating Frame and Glass areas.
*   **Geometric Fallback:** If property set areas are missing, it calculates area physically using 3D geometry (`ifcopenshell.geom`).
*   **Excel Export:** Generates a multi-sheet Excel report on your Desktop containing both summaries and raw data.

---

## 🛠️ Prerequisites & Installation

You need **Python 3.x** installed.

### 1. Install Dependencies
This tool relies on `ifcopenshell` for BIM processing and `pandas` for data handling.

```bash
pip install ifcopenshell pandas numpy openpyxl
