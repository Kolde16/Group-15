import ifcopenshell
import ifcopenshell.util.element
import pandas as pd
import os
import re

ifc_file = r"25-16-D-ARCH.ifc"   # <-- change if needed
model = ifcopenshell.open(ifc_file)

# helpers
_re_mm = re.compile(r"(\d+(\.\d+)?)\s*mm", flags=re.IGNORECASE)
_re_m  = re.compile(r"(\d+(\.\d+)?)\s*m\b", flags=re.IGNORECASE)

def read_prop_value(v):
    """Robustly unwrap Ifc value objects or return raw python value."""
    if v is None:
        return None
    # If it's an IfcValue wrapper with wrappedValue
    try:
        if hasattr(v, "wrappedValue"):
            return v.wrappedValue
    except Exception:
        pass
    return v

def parse_thickness_text(s):
    """Parse thickness from strings like '22mm' or '0.022 m' into meters."""
    if not s:
        return None
    s = str(s)
    m = _re_mm.search(s)
    if m:
        return float(m.group(1)) / 1000.0
    m = _re_m.search(s)
    if m:
        return float(m.group(1))
    # fallback single number heuristic
    m = re.search(r"(\d+(\.\d+)?)", s)
    if m:
        val = float(m.group(1))
        if val > 10:
            return val / 1000.0
        elif val <= 1:
            return val
        else:
            return val / 1000.0
    return None

def get_conductivity_from_material(mat):
    """Try to read thermal conductivity from material property sets (case-insensitive)."""
    if mat is None:
        return None
    # try util.element.get_psets first (works in many exports)
    try:
        psets = ifcopenshell.util.element.get_psets(mat) or {}
    except Exception:
        psets = {}
    # search common pset and generic keys
    for pset_name, props in psets.items():
        for k, v in (props or {}).items():
            kv = (k or "").lower()
            if any(tok in kv for tok in ("thermalconductivity", "thermal conductivity", "lambda", "conductivity")):
                return float(read_prop_value(v))
    # fallback: inspect HasProperties (some IFCs use IfcMaterialProperties directly)
    try:
        if hasattr(mat, "HasProperties"):
            for propdef in getattr(mat, "HasProperties") or []:
                # many material property defs keep values in .Properties or .HasProperties
                entries = []
                if hasattr(propdef, "Properties"):
                    entries += list(getattr(propdef, "Properties") or [])
                if hasattr(propdef, "HasProperties"):
                    entries += list(getattr(propdef, "HasProperties") or [])
                for p in entries:
                    name = getattr(p, "Name", "") or ""
                    if any(tok in name.lower() for tok in ("thermalconductivity", "lambda", "conductivity")):
                        v = read_prop_value(getattr(p, "NominalValue", None) or getattr(p, "LengthValue", None))
                        if v is not None:
                            return float(v)
    except Exception:
        pass
    return None

# defaults
DEFAULT_FRAME_K = 0.2        # W/mK if material k missing
DEFAULT_FRAME_THICK_M = 0.05 # m fallback thickness if nothing found

rows = []

for win in model.by_type("IfcWindow"):
    try:
        gid = getattr(win, "GlobalId", "")
        name = getattr(win, "Name", "")

        # psets on the window
        psets = ifcopenshell.util.element.get_psets(win) or {}

        # basic dims
        ow = getattr(win, "OverallWidth", None)   # likely mm
        oh = getattr(win, "OverallHeight", None)
        # area in Dimensions pset (likely in m2 already)
        pset_area = None
        if "Dimensions" in psets:
            pset_area = read_prop_value(psets["Dimensions"].get("Area"))

        # frame props from psets
        frame_width_raw = None
        rebate_depth_raw = None
        for pset_name, props in (psets or {}).items():
            for k, v in (props or {}).items():
                kl = (k or "").lower()
                if "frame" in kl and "width" in kl:
                    frame_width_raw = read_prop_value(v)
                if ("rebate" in kl and "depth" in kl) or (("frame" in kl) and ("depth" in kl or "thickness" in kl)):
                    rebate_depth_raw = read_prop_value(v)

        # normalize numeric lengths to meters
        def norm_to_m(x):
            if x is None:
                return None
            try:
                xf = float(x)
                # heuristic: if > 10 assume mm, else assume m
                return (xf / 1000.0) if xf > 10 else xf
            except Exception:
                return parse_thickness_text(str(x))

        frame_width_m = norm_to_m(frame_width_raw)
        rebate_depth_m = norm_to_m(rebate_depth_raw)

        # total area (m2)
        total_area_m2 = None
        if ow and oh:
            try:
                total_area_m2 = (float(ow) * float(oh)) / 1e6
            except Exception:
                total_area_m2 = pset_area
        else:
            total_area_m2 = pset_area

        # find materials and their conductivities (glass vs frame)
        glass_mat_name = None
        frame_mat_name = None
        glass_k = None
        frame_k = None
        frame_thickness_from_mat = None

        if hasattr(win, "HasAssociations"):
            for rel in getattr(win, "HasAssociations") or []:
                if not rel or not rel.is_a("IfcRelAssociatesMaterial"):
                    continue
                matdef = getattr(rel, "RelatingMaterial", None)
                if matdef is None:
                    continue
                # IfcMaterialConstituentSet common:
                if matdef.is_a("IfcMaterialConstituentSet"):
                    for cons in getattr(matdef, "MaterialConstituents") or []:
                        mat_ent = getattr(cons, "Material", None)
                        cname = (getattr(mat_ent, "Name", "") if mat_ent else getattr(cons, "Name", "")) or ""
                        c_k = get_conductivity_from_material(mat_ent)
                        # thickness maybe on constituent or in its name
                        thr = None
                        # try known attributes on constituent
                        for attr in ("NominalThickness", "LayerThickness", "MaterialThickness", "Thickness"):
                            if hasattr(cons, attr):
                                thr = getattr(cons, attr)
                                if thr is not None:
                                    thr = norm_to_m(thr)
                                    break
                        # parse from material/constituent name
                        if thr is None:
                            thr = parse_thickness_text(cname)
                        # assign
                        if "glass" in cname.lower() or "glaz" in cname.lower():
                            glass_mat_name = cname
                            if c_k is not None:
                                glass_k = c_k
                        else:
                            # frame
                            if not frame_mat_name:
                                frame_mat_name = cname
                            if frame_k is None and c_k is not None:
                                frame_k = c_k
                            if frame_thickness_from_mat is None and thr is not None:
                                frame_thickness_from_mat = thr
                elif matdef.is_a("IfcMaterial"):
                    mname = getattr(matdef, "Name", "") or ""
                    m_k = get_conductivity_from_material(matdef)
                    if "glass" in mname.lower() or "glaz" in mname.lower():
                        glass_mat_name = mname
                        if m_k is not None:
                            glass_k = m_k
                    else:
                        if not frame_mat_name:
                            frame_mat_name = mname
                        if frame_k is None and m_k is not None:
                            frame_k = m_k
                        # try parse thickness from material name
                        tparse = parse_thickness_text(mname)
                        if frame_thickness_from_mat is None and tparse is not None:
                            frame_thickness_from_mat = tparse

        # final frame thickness: prefer rebate depth (pset), then material-parsed thickness, then fallback
        final_frame_thickness_m = rebate_depth_m or frame_thickness_from_mat or DEFAULT_FRAME_THICK_M

        # final frame conductivity: default if missing
        final_frame_k = frame_k if (frame_k is not None) else DEFAULT_FRAME_K

        # compute U-values (glass_u is conductivity per your rule)
        glass_u = float(glass_k) if glass_k is not None else None
        frame_u = final_frame_k * final_frame_thickness_m if final_frame_thickness_m is not None else None

        # compute areas
        glass_area_m2 = None
        frame_area_m2 = None
        if total_area_m2 is not None and ow and oh:
            if frame_width_m is not None:
                w_m = float(ow) / 1000.0
                h_m = float(oh) / 1000.0
                inner_w = max(0.0, w_m - 2.0 * frame_width_m)
                inner_h = max(0.0, h_m - 2.0 * frame_width_m)
                glass_area_m2 = inner_w * inner_h
                frame_area_m2 = max(0.0, total_area_m2 - glass_area_m2)
            else:
                # assume whole is glazing if no frame width
                glass_area_m2 = total_area_m2
                frame_area_m2 = 0.0

        # compute overall window U (area-weighted)
        window_u = None
        try:
            if glass_u is not None and frame_u is not None and glass_area_m2 is not None and frame_area_m2 is not None:
                denom = (glass_area_m2 + frame_area_m2)
                if denom > 0:
                    window_u = (glass_u * glass_area_m2 + frame_u * frame_area_m2) / denom
        except Exception:
            window_u = None

        rows.append({
            "GlobalId": gid,
            "Name": name,
            "Width_mm": ow,
            "Height_mm": oh,
            "TotalArea_m2": total_area_m2,
            "FrameWidth_m": frame_width_m,
            "RebateDepth_m": rebate_depth_m,
            "FrameThickness_mat_m": frame_thickness_from_mat,
            "FrameThickness_final_m": final_frame_thickness_m,
            "FrameMaterial": frame_mat_name,
            "Frame_k_W_mK": final_frame_k,
            "Frame_U_W_m2K": frame_u,
            "FrameArea_m2": frame_area_m2,
            "GlassMaterial": glass_mat_name,
            "Glass_k_W_mK": glass_k,
            "Glass_U_W_m2K": glass_u,
            "GlassArea_m2": glass_area_m2,
            "Window_U_W_m2K": window_u
        })

    except Exception as e:
        print(f"ERROR processing window (GlobalId={getattr(win,'GlobalId','?')}): {e}")

# Save to Excel
df = pd.DataFrame(rows)
base = os.path.splitext(os.path.basename(ifc_file))[0]
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
out_path = os.path.join(desktop, f"{base}_windows_uvalues.xlsx")
df.to_excel(out_path, index=False)
print("✅ Done — Excel saved to:", out_path)
