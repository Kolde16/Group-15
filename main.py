import ifcopenshell
import ifcopenshell.util.element
import ifcopenshell.util.placement
import ifcopenshell.geom
import pandas as pd
import numpy as np
import os
import re
import sys

# --- CONFIGURATION ---
IFC_FILENAME = "25-16-D-ARCH.ifc"

# Window Defaults
DEFAULTS = {
    "FRAME_K": 0.17,       # W/mK
    "FRAME_THICK": 0.07,   # m
    "GLASS_U": 1.2         # W/m2K
}

# --- 1. SETUP ENV & GEOMETRY ---
try:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    SCRIPT_DIR = os.getcwd()

IFC_PATH = os.path.join(SCRIPT_DIR, IFC_FILENAME)

try:
    settings = ifcopenshell.geom.settings()
    settings.set(settings.USE_WORLD_COORDS, True)
    GEOM_AVAILABLE = True
except:
    print("⚠️ Warning: ifcopenshell.geom not available. Geometric calc for roofs will fail.")
    GEOM_AVAILABLE = False

# --- 2. SHARED UTILITY FUNCTIONS ---

def get_real_value(v):
    if hasattr(v, "wrappedValue"): return v.wrappedValue
    return v

def is_valid_numeric_string(s):
    s = str(s).strip().lower()
    allowed = set("0123456789.me+- ")
    if not set(s).issubset(allowed): return False
    return True

def clean_numeric(val):
    if val is None: return None
    s = str(val).strip().lower()
    if not is_valid_numeric_string(s): return None
    try:
        if "mm" in s:
            num = re.search(r"(\d+(\.\d+)?)", s)
            return float(num.group(1)) / 1000.0 if num else None
        f = float(val)
        if f > 20: return f / 1000.0
        return f
    except ValueError: return None

def get_properties_merged(element):
    all_props = {}
    try:
        if hasattr(element, "IsTypedBy"):
            for rel in element.IsTypedBy:
                type_entity = rel.RelatingType
                type_psets = ifcopenshell.util.element.get_psets(type_entity)
                for pset_name, props in type_psets.items():
                    if pset_name not in all_props: all_props[pset_name] = {}
                    all_props[pset_name].update(props)
    except: pass
    inst_psets = ifcopenshell.util.element.get_psets(element)
    for pset_name, props in inst_psets.items():
        if pset_name not in all_props: all_props[pset_name] = {}
        all_props[pset_name].update(props)
    return all_props

def find_prop_value(merged_props, keys_to_find, precise_pset=None):
    # Exact Match
    for pset_name, props in merged_props.items():
        if precise_pset and precise_pset.lower() not in pset_name.lower(): continue
        for k, v in props.items():
            if k.lower() in [key.lower() for key in keys_to_find]:
                return get_real_value(v)
    # Partial Match
    for pset_name, props in merged_props.items():
        if precise_pset and precise_pset.lower() not in pset_name.lower(): continue
        for k, v in props.items():
            for key in keys_to_find:
                if key.lower() in k.lower(): return get_real_value(v)
    return None

def get_geo_loc(element):
    try:
        matrix = ifcopenshell.util.placement.get_local_placement(element.ObjectPlacement)
        x, y, z = matrix[0][3], matrix[1][3], matrix[2][3]
        return f"{round(x,3)}_{round(y,3)}_{round(z,3)}"
    except: return "Unknown"

# --- 3. MODULES ---

def process_windows(model):
    print("   ...processing Windows")
    rows = []
    
    def get_window_materials(win):
        data = {"frame_mat": None, "frame_k": None, "frame_thick_mat": None, "glass_mat": None, "glass_k": None}
        if not hasattr(win, "HasAssociations"): return data
        materials = []
        try:
            for rel in getattr(win, "HasAssociations", []):
                if rel.is_a("IfcRelAssociatesMaterial"):
                    mat = rel.RelatingMaterial
                    if mat.is_a("IfcMaterialConstituentSet"):
                        materials.extend([c.Material for c in getattr(mat, "MaterialConstituents", []) if c.Material])
                    elif mat.is_a("IfcMaterialList"):
                        materials.extend(getattr(mat, "Materials", []) or [])
                    elif mat.is_a("IfcMaterial"):
                        materials.append(mat)
        except: pass

        for m in materials:
            name = (getattr(m, "Name", "") or "").lower()
            k = None
            try:
                psets = ifcopenshell.util.element.get_psets(m) or {}
                for props in psets.values():
                    for pk, pv in props.items():
                        if any(x in pk.lower() for x in ["thermalconductivity", "lambda"]):
                            k = float(get_real_value(pv))
            except: pass

            if "glass" in name or "glaz" in name:
                if not data["glass_mat"]: data["glass_mat"] = getattr(m, "Name", "Glass")
                if k: data["glass_k"] = k
            else:
                if not data["frame_mat"]: data["frame_mat"] = getattr(m, "Name", "Frame")
                if k: data["frame_k"] = k
                t = clean_numeric(getattr(m, "Name", ""))
                if t: data["frame_thick_mat"] = t
        return data

    for win in model.by_type("IfcWindow"):
        try:
            props = get_properties_merged(win)
            ow = getattr(win, "OverallWidth", None)
            oh = getattr(win, "OverallHeight", None)
            w_m = clean_numeric(ow)
            h_m = clean_numeric(oh)
            frame_w = clean_numeric(find_prop_value(props, ["FrameWidth", "Frame Width"]))
            
            area_total = (w_m * h_m) if (w_m and h_m) else clean_numeric(find_prop_value(props, ["Area"], "Dimensions"))
            mat_data = get_window_materials(win)
            
            f_thick = mat_data['frame_thick_mat'] or DEFAULTS['FRAME_THICK']
            f_k = mat_data['frame_k'] if mat_data['frame_k'] else DEFAULTS['FRAME_K']
            f_u = (f_k / f_thick) if (f_thick and f_thick > 0) else None
            g_u = mat_data['glass_k'] if (mat_data['glass_k'] and 0.1 < mat_data['glass_k'] < 3.5) else DEFAULTS['GLASS_U']
            
            w_u_total = None
            if area_total and f_u and g_u:
                if frame_w and w_m and h_m:
                    inner_w = max(0.0, w_m - 2*frame_w)
                    inner_h = max(0.0, h_m - 2*frame_w)
                    g_area = inner_w * inner_h
                    f_area = max(0.0, area_total - g_area)
                else:
                    f_area = area_total * 0.15
                    g_area = area_total * 0.85
                w_u_total = ((g_u * g_area) + (f_u * f_area)) / area_total

            rows.append({
                "GlobalId": win.GlobalId,
                "Name": win.Name,
                "Type_Name": "Window",
                "Category": "Opening",
                "FrameMat": mat_data['frame_mat'],
                "GlassMat": mat_data['glass_mat'],
                "GeoLocation": get_geo_loc(win),
                "Width_mm": (w_m * 1000) if w_m else None,
                "Height_mm": (h_m * 1000) if h_m else None,
                "Area_m2": area_total,
                "U_Value": w_u_total
            })
        except Exception: pass
    return pd.DataFrame(rows)

def process_walls(model):
    print("   ...processing Walls")
    rows = []
    for wall in model.by_type("IfcWall"):
        try:
            props = get_properties_merged(wall)
            u_val = clean_numeric(find_prop_value(props, ["ThermalTransmittance", "Heat Transfer Coefficient (U)", "U Value"]))
            width_m = clean_numeric(find_prop_value(props, ["Width", "Thickness"]))
            area_m = clean_numeric(find_prop_value(props, ["Area", "NetSideArea"], "Dimensions"))
            if not area_m: area_m = clean_numeric(find_prop_value(props, ["NetSideArea", "Area"]))
            
            pos = "Internal"
            is_ext = find_prop_value(props, ["IsExternal"])
            func = find_prop_value(props, ["Function"])
            if (str(is_ext).lower() in ['true', '1', 'yes']) or (func and "ext" in str(func).lower()): pos = "External"
            
            t_name = "Unknown"
            if hasattr(wall, "IsTypedBy") and wall.IsTypedBy: t_name = wall.IsTypedBy[0].RelatingType.Name

            rows.append({
                "GlobalId": wall.GlobalId,
                "Name": wall.Name,
                "Type_Name": t_name,
                "Category": "Wall",
                "Position": pos,
                "GeoLocation": get_geo_loc(wall),
                "Thickness_mm": (width_m * 1000) if width_m else None,
                "Area_m2": area_m,
                "U_Value": u_val
            })
        except Exception: pass
    return pd.DataFrame(rows)

def calculate_geom_area(elem):
    if not GEOM_AVAILABLE: return None
    try:
        shape = ifcopenshell.geom.create_shape(settings, elem)
        verts = np.array(shape.geometry.verts).reshape(-1, 3)
        faces = np.array(shape.geometry.faces).reshape(-1, 3)
        total_up = 0.0
        for face in faces:
            p1, p2, p3 = verts[face[0]], verts[face[1]], verts[face[2]]
            cross = np.cross(p2 - p1, p3 - p1)
            area = 0.5 * np.linalg.norm(cross)
            if area > 0:
                if (cross / np.linalg.norm(cross))[2] > 0: total_up += area
        return total_up
    except: return None

def process_slabs(model):
    print("   ...processing Slabs & Roofs")
    rows = []
    for elem in model.by_type("IfcSlab") + model.by_type("IfcRoof"):
        try:
            props = get_properties_merged(elem)
            ptype = str(getattr(elem, "PredefinedType", "")).upper()
            name = (elem.Name or "").lower()
            is_ext = find_prop_value(props, ["IsExternal"])
            is_ext_bool = str(is_ext).lower() in ['true', '1', 'yes']
            
            cat, pos = "Floor", "Internal"
            if ptype == "ROOF" or "roof" in name: cat, pos = "Roof", "External"
            elif ptype == "BASESLAB" or "found" in name: cat, pos = "Foundation", "External"
            elif is_ext_bool: cat, pos = "Slab", "External"

            u_val = clean_numeric(find_prop_value(props, ["ThermalTransmittance", "Heat Transfer Coefficient (U)", "U Value"]))
            thick_m = clean_numeric(find_prop_value(props, ["Thickness", "Width", "Depth"]))
            area_m = clean_numeric(find_prop_value(props, ["Area", "NetArea"], "Dimensions"))
            if not area_m: area_m = clean_numeric(find_prop_value(props, ["Area", "NetArea"]))
            if not area_m and thick_m:
                vol = clean_numeric(find_prop_value(props, ["NetVolume", "Volume"]))
                if vol: area_m = vol / thick_m
            if not area_m: area_m = calculate_geom_area(elem)

            t_name = "Unknown"
            if hasattr(elem, "IsTypedBy") and elem.IsTypedBy: t_name = elem.IsTypedBy[0].RelatingType.Name

            rows.append({
                "GlobalId": elem.GlobalId,
                "Name": elem.Name,
                "Type_Name": t_name,
                "Category": cat,
                "Position": pos,
                "GeoLocation": get_geo_loc(elem),
                "Thickness_mm": (thick_m * 1000) if thick_m else None,
                "Area_m2": area_m,
                "U_Value": u_val
            })
        except Exception: pass
    return pd.DataFrame(rows)

# --- 4. MAIN EXECUTION ---

def run_main():
    if not os.path.exists(IFC_PATH):
        print(f"❌ ERROR: File not found: {IFC_PATH}")
        sys.exit(1)

    print(f"Loading {IFC_FILENAME}...")
    model = ifcopenshell.open(IFC_PATH)
    
    # 1. Process
    df_win = process_windows(model)
    df_wall = process_walls(model)
    df_slab = process_slabs(model)
    
    # 2. Cleanup
    def clean_df(df, subset_cols):
        if df.empty: return df
        df = df.dropna(subset=subset_cols)
        df = df.drop_duplicates(subset=subset_cols, keep='first')
        for c in ["U_Value", "Area_m2"]:
            if c in df.columns: df[c] = df[c].round(3)
        return df

    df_win = clean_df(df_win, ["GeoLocation", "Width_mm", "Height_mm"])
    df_wall = clean_df(df_wall, ["GeoLocation", "Thickness_mm", "Area_m2"])
    df_slab = clean_df(df_slab, ["GeoLocation", "Thickness_mm", "Area_m2"])

    # 3. Create Specific Summaries
    # FIX: Added dropna=False to ensure rows with missing U-values still show up
    def create_summary(df, group_cols):
        if df.empty: return pd.DataFrame()
        valid = [c for c in group_cols if c in df.columns]
        return df.groupby(valid, dropna=False).agg(
            Count=('GlobalId', 'count'),
            Total_Area=('Area_m2', 'sum'),
            Avg_U=('U_Value', 'mean')
        ).reset_index()

    summ_win = create_summary(df_win, ['Type_Name', 'FrameMat', 'GlassMat', 'Width_mm', 'Height_mm'])
    summ_wall = create_summary(df_wall, ['Category', 'Position', 'Type_Name', 'Thickness_mm', 'U_Value'])
    summ_slab = create_summary(df_slab, ['Category', 'Position', 'Type_Name', 'Thickness_mm', 'U_Value'])

    # 4. Create MASTER SUMMARY
    print("   ...generating Master Summary")
    common_cols = ["GlobalId", "Name", "Category", "Type_Name", "Position", "Area_m2", "U_Value"]
    if not df_win.empty: df_win["Position"] = "External"
    
    master_frames = []
    if not df_win.empty: master_frames.append(df_win[common_cols])
    if not df_wall.empty: master_frames.append(df_wall[common_cols])
    if not df_slab.empty: master_frames.append(df_slab[common_cols])
    
    if master_frames:
        df_master_raw = pd.concat(master_frames, ignore_index=True)
        
        def weighted_avg(x):
            # Calculate average only on valid numbers
            valid_rows = x.dropna(subset=['U_Value'])
            if valid_rows.empty: return None
            total_valid_area = valid_rows['Area_m2'].sum()
            if total_valid_area > 0:
                return np.average(valid_rows['U_Value'], weights=valid_rows['Area_m2'])
            else:
                return valid_rows['U_Value'].mean()

        # Added dropna=False here too just in case
        df_master_summ = df_master_raw.groupby(['Category', 'Position', 'Type_Name'], dropna=False).apply(
            lambda x: pd.Series({
                'Count': len(x),
                'Total_Area': x['Area_m2'].sum(),
                'Weighted_Avg_U': weighted_avg(x)
            })
        ).reset_index().sort_values(["Category", "Position"])
        
        df_master_summ['Weighted_Avg_U'] = df_master_summ['Weighted_Avg_U'].apply(lambda x: round(x, 3) if pd.notnull(x) else None)
        df_master_summ['Total_Area'] = df_master_summ['Total_Area'].round(3)

    else:
        df_master_summ = pd.DataFrame()

    # 5. Export
    out_path = os.path.join(os.path.expanduser("~"), "Desktop", f"{os.path.splitext(IFC_FILENAME)[0]}_THERMAL_REPORT.xlsx")
    
    try:
        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            df_master_summ.to_excel(writer, sheet_name='MASTER SUMMARY', index=False)
            
            if not summ_wall.empty: summ_wall.to_excel(writer, sheet_name='Walls Summary', index=False)
            if not summ_slab.empty: summ_slab.to_excel(writer, sheet_name='Slabs Summary', index=False)
            if not summ_win.empty: summ_win.to_excel(writer, sheet_name='Windows Summary', index=False)
            
            if not df_wall.empty: df_wall.to_excel(writer, sheet_name='Walls Data', index=False)
            if not df_slab.empty: df_slab.to_excel(writer, sheet_name='Slabs Data', index=False)
            if not df_win.empty: df_win.to_excel(writer, sheet_name='Windows Data', index=False)

        print(f"\n✅ SUCCESS! Report generated at: {out_path}")

    except PermissionError:
        print("\n❌ ERROR: Please close the Excel file and run again.")

if __name__ == "__main__":
    run_main()
