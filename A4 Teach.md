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

This tool is designed to run locally on your computer. It is optimized for the **Spyder IDE** (typically installed with Anaconda).

### 1. Install Dependencies
This tool relies on `ifcopenshell` for BIM processing and `pandas` for data handling.

```bash
pip install ifcopenshell pandas numpy openpyxl
```

*> **Note:** If `pip install ifcopenshell` fails (as it depends on C++ libraries), it is recommended to install via Conda: `conda install -c conda-forge ifcopenshell`*

### 2. Folder Setup
The script is programmed to look for the IFC file in the **exact same folder** as the Python file.

1.  Create a new folder on your computer (e.g., `Documents/Thermal_Tool`).
2.  Save the code as `main.py` inside that folder.
3.  **Copy your `.ifc` file into that exact same folder.**

**Your folder structure must look like this:**
```text
📂 Thermal_Tool
 ├── 📄 main.py
 └── 📄 25-16-D-ARCH.ifc
```

### 3. Configuration
Open `main.py` and set your target file near the top of the script:

```python
# --- CONFIGURATION ---
IFC_FILENAME = "25-16-D-ARCH.ifc"  # <--- CHANGE THIS to your filename
```

You can also adjust the default assumptions for windows if manufacturer data is missing:
```python
DEFAULTS = {
    "FRAME_K": 0.17,       # Frame thermal conductivity (W/mK)
    "FRAME_THICK": 0.07,   # Frame thickness (m)
    "GLASS_U": 1.2         # Glass center-of-glazing U-value
}
```

---

## 🧠 Technical Walkthrough

This section explains the code pipeline step-by-step for developers.

### Step 1: Environment & Geometry Setup
The script begins by initializing the `ifcopenshell` geometry engine. This is required to calculate areas for complex roof shapes where explicit property values might be missing.

```python
try:
    settings = ifcopenshell.geom.settings()
    settings.set(settings.USE_WORLD_COORDS, True)
    GEOM_AVAILABLE = True
except:
    print("⚠️ Warning: ifcopenshell.geom not available...")
```

### Step 2: The "Property Hunt" (`get_properties_merged`)
IFC data is hierarchical. A window might have a U-value defined in its **Type** (shared by all windows of that style) or its **Instance** (specific to that one window).
The script flattens these into a single dictionary, prioritizing **Instance** properties if duplicates exist.

```python
def get_properties_merged(element):
    all_props = {}
    # 1. Grab Type Properties (e.g., IfcWindowStyle)
    if hasattr(element, "IsTypedBy"):
        # ... logic to get type psets ...
    
    # 2. Grab Instance Properties (Overwrites Type)
    inst_psets = ifcopenshell.util.element.get_psets(element)
    return all_props
```

### Step 3: Data Normalization (`clean_numeric`)
Architects input data inconsistently. One might write "150mm", another "0.15m", and another just "150". This function uses Regex to strip units and normalize everything to **meters**.

```python
def clean_numeric(val):
    # Detects "mm" in string or values > 20 (assumed mm)
    if "mm" in s:
        return float(num.group(1)) / 1000.0
    if f > 20: return f / 1000.0 
    return f
```

### Step 4: Element Processing
The script iterates through specific IFC entities. Each has unique logic:

#### A. Windows (Composite Logic)
Windows require separating Frame area from Glass area. The script scans `IfcRelAssociatesMaterial` to find specific materials.
$$ U_{total} = \frac{(U_{glass} \times A_{glass}) + (U_{frame} \times A_{frame})}{A_{total}} $$

#### B. Slabs & Roofs (Geometric Logic)
The script attempts to classify slabs (Floor vs Roof vs Foundation) based on their `PredefinedType` and Name. If `NetArea` is missing in the properties, it triggers the geometry engine:
```python
# Calculates the area of all mesh faces pointing 'UP' (Z > 0)
def calculate_geom_area(elem):
    # ... create shape ...
    for face in faces:
        if (cross / np.linalg.norm(cross))[2] > 0: 
            total_up += area
    return total_up
```

### Step 5: Data Cleaning & Aggregation
Once the raw data is in a Pandas DataFrame, the script removes rows with invalid geometry (`NaN`) and groups them.

Crucially, it calculates an **Area-Weighted Average** for the U-Values. A small window shouldn't skew the average as much as a large curtain wall.

```python
def weighted_avg(x):
    # Numpy average using 'Area_m2' as the weight
    return np.average(valid_rows['U_Value'], weights=valid_rows['Area_m2'])
```

### Step 6: Excel Export
Finally, the script writes the data to an Excel file with multiple sheets:
1.  **MASTER SUMMARY:** High-level totals.
2.  **Category Summaries:** Grouped by Type.
3.  **Raw Data:** The full dataset for debugging.

---

## ❓ Design Decisions

### Why are Doors excluded?
The script currently ignores `IfcDoor`.
*   **Reasoning:** Doors are complex hybrids (part opaque panel, part glazing, part frame). Applying standard "Window Logic" to a solid wood door would yield incorrect thermal data.
*   **Future:** A dedicated `process_doors` function is needed.

### Why use `try/except` for properties?
IFC files vary wildly between software (Revit, ArchiCAD, Tekla). The script uses broad partial string matching (finding "Width" inside "Frame Width") and `try/except` blocks to ensure the script continues running even if one specific element is corrupt or non-standard.

---

## 🏃 Usage

Run the script from your terminal:

```bash
python main.py
```

The report will be generated on your **Desktop**.
