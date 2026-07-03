"""
Process Engineering Automation — Streamlit app
==============================================
Two tabs:
  1. Equipment Selection  — reactor / filter / dryer selector (original logic, cleaned).
  2. PFD Generator        — full Python port of the 5 MATLAB steps + VBA formatter.
                            Report (PDF / DOCX / pasted text) + optional material sheet
                            -> numbered steps, conditional decisions, material & qty
                            assignment, cumulative Min/Max volume ledger, and a
                            formatted flowchart Excel. No MATLAB / VBA required.
"""

import io

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

import pfd_pipeline as pfd
import pfd_scaleup as scaleup
import report_parser as rparse
import pfd_formula_writer as fwriter

st.set_page_config(page_title="Process Engineering Automation", layout="wide")


# ==========================================================================
#  Equipment-selection helpers  (ported/cleaned from the original code.py)
# ==========================================================================
def load_reactor_data(filepath):
    df = pd.read_excel(filepath)
    df.columns = df.columns.str.strip().str.lower()
    rename_map = {
        "vessel id": "reactor id",
        "min sensing volume": "min sensing",
        "min stirring volume": "min stirring",
        "capacity": "max volume",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    if "utilities" not in df.columns:
        st.error("'utilities' column not found in the uploaded Excel.")
        return pd.DataFrame()

    df["moc"] = df["moc"].str.upper().replace({"ALL GLASS": "GLR"})
    df["materials"] = df["moc"].apply(lambda x: [m.strip() for m in x.split("/")])
    df["thermal options"] = df["utilities"].astype(str).apply(
        lambda x: [t.strip().upper() for t in x.split(",")])
    df["agitator"] = df["agitator"].astype(str).str.upper()
    return df[["reactor id", "min sensing", "min stirring", "max volume",
               "materials", "thermal options", "agitator"]]


def load_filter_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip().str.lower()
    return df


def load_dryer_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip().str.lower()
    return df.rename(columns={"dryer id": "equipment id"})


def _allowed_moc(ph, user_input):
    if ph == "basic":
        return ["SSR", "HAR", "HALAR"]
    if ph == "acidic":
        return ["GLR", "HALAR", "HAR"]
    if ph == "neutral":
        return ["GLR", "SSR", "HAR", "HALAR"]
    if ph == "coupon":
        mat = user_input["coupon_materials"][0].strip().upper()
        if user_input["corrosion_rate"] < 0.1:
            return [mat]
        st.error("Corrosion rate too high for this material.")
        return []
    return []


def filter_reactors(df, user_input, first_step_vol, total_vol):
    df = df[(df["min sensing"] <= first_step_vol) & (df["min stirring"] <= first_step_vol)]
    vol_limit = 0.7 if user_input["process_type"] in ["distillation", "reaction", "pressurized"] else 0.95
    df = df[df["max volume"] * vol_limit >= total_vol]

    allowed = _allowed_moc(user_input["ph_condition"], user_input)
    if not allowed:
        return pd.DataFrame()
    # reactor MOC uses GLR/SSR/HAR (no HALAR); intersect with reactor vocabulary
    reactor_allowed = [m for m in allowed if m in {"GLR", "SSR", "HAR"}]
    df = df[df["materials"].apply(lambda mats: any(m in mats for m in reactor_allowed))]

    temp = user_input["temperature"]
    if 10 <= temp <= 20:
        thermal = ["CHB"]
    elif 20 < temp <= 35:
        thermal = ["CT"]
    elif 35 < temp <= 90:
        thermal = ["HW"]
    else:
        thermal = ["LPS", "HOT OIL", "EJECTION CONDENSATE"]
    df = df[df["thermal options"].apply(lambda opts: any(t in opts for t in thermal))]

    preferred = []
    if user_input["reaction_nature"] == "homogeneous":
        preferred = ["PROPELLOR", "PBT", "RCI", "ANCHOR", "CBRT"]
    elif user_input["reaction_nature"] == "heterogeneous":
        st_ = user_input["reaction_subtype"]
        if st_ == "biphasic":
            preferred = ["PROPELLOR", "PBT", "CBRT", "RCI"]
        elif st_ == "solid-liquid":
            preferred = ["PROPELLOR", "PBT", "CBRT", "RCI", "ANCHOR"]
        elif st_ == "gas-liquid":
            preferred = ["RUSTON", "DISC"]

    df = df.copy()
    df["Preference Match"] = df["agitator"].apply(
        lambda a: "yes" if any(p in a for p in preferred) else "warning")
    return df


def filter_filters(df, user_input, filter_types_required):
    allowed = _allowed_moc(user_input["ph_condition"], user_input)
    if not allowed:
        return pd.DataFrame()
    df = df[df["moc"].astype(str).str.upper().isin(allowed)]

    volume_litres = (user_input["mass"] / user_input["bulk_density"] * 1000
                     if user_input["bulk_density"] > 0 else 0)
    st.write(f"Volume required (L): {volume_litres:.2f}")

    if "cake capacity" not in df.columns:
        st.error("'cake capacity' column not found in the uploaded Excel.")
        return pd.DataFrame()
    df = df[df["cake capacity"] * 0.9 >= volume_litres]

    if not filter_types_required:
        st.warning("No filter type matched the selected filter property.")
        return pd.DataFrame()
    if "filter type" not in df.columns:
        st.error("'filter type' column not found in the uploaded Excel.")
        return pd.DataFrame()
    df = df.copy()
    df["filter type"] = df["filter type"].astype(str).str.upper()
    return df[df["filter type"].apply(lambda x: any(f in x for f in filter_types_required))]


def filter_dryers(df, user_input):
    allowed = _allowed_moc(user_input["ph_condition"], user_input)
    if not allowed:
        return pd.DataFrame()
    df = df[df["moc"].astype(str).str.upper().isin(allowed)]
    st.write(f"Volume required (L): {user_input['volume']:.2f}")
    if "capacity" not in df.columns:
        st.error("'capacity' column not found in the uploaded Excel.")
        return pd.DataFrame()
    return df[df["capacity"] * 0.9 >= user_input["volume"]]


def collect_unit_operation(unit_op_id):
    steps, total_volume, first_step_volume = [], 0.0, None
    st.subheader("Add Steps to Unit Operation")
    n = st.number_input("How many steps in this unit operation?",
                        min_value=1, max_value=30, value=1, key=f"nsteps_{unit_op_id}")
    for i in range(1, int(n) + 1):
        st.markdown(f"### Step {i}")
        c1, c2, c3 = st.columns(3)
        operation = c1.selectbox("Operation", ["charge", "addition"], key=f"op_{unit_op_id}_{i}")
        material = c2.selectbox("Material", ["reagent 1", "reagent 2", "reagent 3", "KSM", "solvent"],
                                key=f"mat_{unit_op_id}_{i}")
        volume = c3.number_input("Volume (L)", min_value=0.0, key=f"vol_{unit_op_id}_{i}")
        actual = volume
        if material == "KSM":
            pct = st.number_input("KSM %", min_value=0.0, max_value=100.0, key=f"ksm_{unit_op_id}_{i}")
            if pct > 0:
                actual = volume / (pct / 100)
        if first_step_volume is None:
            first_step_volume = actual
        total_volume += actual
        steps.append({"unit_op": unit_op_id, "step": i, "operation": operation,
                      "material": material, "input_volume": volume,
                      "actual_volume": actual, "accumulated_volume": total_volume})
    return first_step_volume, total_volume, steps


def export_steps_to_excel(steps_by_unitop):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        rows = []
        for uid, (steps, reactor) in enumerate(steps_by_unitop, start=1):
            for s in steps:
                rows.append({
                    "Unit Operation": uid, "Operation": s["operation"],
                    "Material": s["material"], "Volume Added (L)": s["actual_volume"],
                    "Accumulated Volume (L)": s["accumulated_volume"], "Equipment": reactor})
        df = pd.DataFrame(rows)
        df.to_excel(writer, index=False, sheet_name="Steps")
        ws = writer.sheets["Steps"]
        for col in ws.columns:
            width = max((len(str(c.value)) if c.value else 0) for c in col) + 2
            ws.column_dimensions[get_column_letter(col[0].column)].width = width
    buffer.seek(0)
    return buffer


# ==========================================================================
#  TAB 1 — Equipment Selection
# ==========================================================================
def tab_equipment():
    st.title("Process Engineering Automation")

    if "selections" not in st.session_state:
        st.session_state.selections = []

    with st.sidebar:
        st.header("Unit Operation Steps")
        if st.session_state.selections:
            for i, (log, sel) in enumerate(st.session_state.selections):
                st.markdown(f"### Step {i + 1}")
                for step in log:
                    st.markdown(f"- **Op:** {step.get('operation','N/A')} | "
                                f"**Mat:** {step.get('material','N/A')} | "
                                f"**Vol:** {step.get('actual_volume',0)} L")
                st.markdown(f"**Selected Equipment:** {sel}")
                if st.button(f"Remove Step {i + 1}", key=f"remove_{i}"):
                    st.session_state.selections.pop(i)
                    st.rerun()
        else:
            st.info("No unit operations added yet.")

    uploaded_file = st.file_uploader("Upload reactor database", type="xlsx")
    if not uploaded_file:
        st.info("Upload the reactor database to start.")
        return

    df = load_reactor_data(uploaded_file)
    if df.empty:
        return

    batch_id = len(st.session_state.selections) + 1
    st.header("Enter Process Conditions")
    st.markdown(f"## Unit Operation {batch_id}")

    unit_op_type = st.selectbox("Unit operation type",
                                ["reaction", "distillation", "pressurized",
                                 "extraction/workup", "filtration", "drying"],
                                key=f"unit_type_{batch_id}")
    ph_condition = st.selectbox("pH condition", ["basic", "acidic", "neutral", "coupon"],
                                key=f"ph_{batch_id}")
    corrosion_rate, coupon_materials = 0.0, []
    if ph_condition == "coupon":
        corrosion_rate = st.number_input("Corrosion rate (mm/year)", min_value=0.0, key=f"cr_{batch_id}")
        coupon_materials = [st.text_input("Coupon material", key=f"cm_{batch_id}").upper()]
    temperature = st.number_input("Process temperature (°C)", key=f"temp_{batch_id}")

    if unit_op_type not in ["filtration", "drying"]:
        reaction_nature = st.selectbox("Nature of reaction", ["none", "homogeneous", "heterogeneous"],
                                       key=f"rn_{batch_id}")
        reaction_subtype = None
        if reaction_nature == "heterogeneous":
            reaction_subtype = st.selectbox("Subtype", ["biphasic", "solid-liquid", "gas-liquid"],
                                            key=f"rs_{batch_id}")
        st.markdown("---")
        first_vol, total_vol, log = collect_unit_operation(batch_id)
        if st.button(f"Submit Unit Operation {batch_id}", key=f"submit_{batch_id}"):
            ui = {"process_type": unit_op_type, "ph_condition": ph_condition,
                  "corrosion_rate": corrosion_rate, "coupon_materials": coupon_materials,
                  "temperature": temperature, "reaction_nature": reaction_nature,
                  "reaction_subtype": reaction_subtype}
            matched = filter_reactors(df.copy(), ui, first_vol, total_vol)
            if not matched.empty:
                st.success(f"Reactors matching Unit Operation {batch_id}")
                view = matched[["reactor id", "min sensing", "min stirring",
                                "max volume", "agitator", "Preference Match"]]
                st.dataframe(view.style.applymap(
                    lambda v: "background-color:#d4edda" if v == "yes"
                    else ("background-color:#fff3cd" if v == "warning" else ""),
                    subset=["Preference Match"]))
                sel = st.selectbox("Select one reactor:", view["reactor id"].tolist(),
                                   key=f"sel_reactor_{batch_id}")
                st.session_state.selections.append((log, sel))
            else:
                st.warning("No matching reactors found.")

    elif unit_op_type == "filtration":
        f = st.file_uploader(f"Upload Filter Database (UO {batch_id})", type=["xlsx"],
                             key=f"upf_{batch_id}")
        if f:
            fdf = load_filter_data(f)
            mass = st.number_input("Mass (kg)", min_value=0.0, key=f"mass_{batch_id}")
            bd = st.number_input("Bulk density (kg/m³)", min_value=0.0, key=f"bd_{batch_id}")
            prop = st.selectbox("Filter property",
                                ["specific cake resistance (m/kg)", "rate of cake buildup", "settling rate"],
                                key=f"fp_{batch_id}")
            req = []
            if prop == "specific cake resistance (m/kg)":
                v = st.number_input("Specific cake resistance (m/kg)", min_value=0.0, key=f"scr_{batch_id}")
                if 1e7 <= v < 1e8:
                    req = ["CENTRIFUGE", "NUTSCHE"]
                elif 1e8 <= v < 1e10:
                    req = ["CENTRIFUGE", "ANFD", "RPF", "VNF"]
                elif v >= 1e10:
                    req = ["CENTRIFUGE", "NUTSCHE"]
            elif prop == "rate of cake buildup":
                unit = st.selectbox("Unit", ["cm/sec", "cm/min", "cm/hr"], key=f"bu_{batch_id}")
                v = st.number_input(f"Rate ({unit})", min_value=0.0, key=f"buv_{batch_id}")
                if unit == "cm/sec" and 0.1 <= v <= 10:
                    req = ["CENTRIFUGE", "NUTSCHE"]
                elif unit == "cm/min" and 0.1 <= v <= 10:
                    req = ["CENTRIFUGE", "ANFD", "RPF"]
                elif unit == "cm/hr" and 0.1 <= v <= 10:
                    req = ["ANFD"]
            else:
                v = st.number_input("Settling rate (cm/sec)", min_value=0.0, key=f"sr_{batch_id}")
                if v > 5:
                    req = ["CENTRIFUGE", "NUTSCHE"]
                elif 0.1 <= v <= 5:
                    req = ["ANFD", "RPF"]
                elif v < 0.1:
                    req = ["ANFD"]
            if st.button(f"Submit Filtration {batch_id}", key=f"subf_{batch_id}"):
                ui = {"ph_condition": ph_condition, "corrosion_rate": corrosion_rate,
                      "coupon_materials": coupon_materials, "temperature": temperature,
                      "bulk_density": bd, "mass": mass}
                matched = filter_filters(fdf.copy(), ui, req)
                if not matched.empty:
                    st.success("Matching filters found")
                    st.dataframe(matched)
                    id_col = next((c for c in matched.columns
                                   if c.strip().lower() in ["equipment id", "filter id", "id"]), None)
                    opts = matched[id_col].astype(str).tolist() if id_col else matched.index.astype(str).tolist()
                    sel = st.selectbox("Select filter:", opts, key=f"self_{batch_id}")
                    vol = mass / bd * 1000 if bd > 0 else 0
                    st.session_state.selections.append(
                        ([{"unit_op": batch_id, "operation": "filtration", "material": "N/A",
                           "input_volume": 0, "actual_volume": vol, "accumulated_volume": vol}], sel))
                else:
                    st.warning("No matching filters found.")

    elif unit_op_type == "drying":
        d = st.file_uploader(f"Upload Dryer Database (UO {batch_id})", type=["xlsx"], key=f"upd_{batch_id}")
        if d:
            ddf = load_dryer_data(d)
            vol = st.number_input("Volume (L)", min_value=0.0, key=f"vd_{batch_id}")
            if st.button(f"Submit Drying {batch_id}", key=f"subd_{batch_id}"):
                ui = {"ph_condition": ph_condition, "corrosion_rate": corrosion_rate,
                      "coupon_materials": coupon_materials, "temperature": temperature, "volume": vol}
                matched = filter_dryers(ddf.copy(), ui)
                if not matched.empty:
                    st.success("Matching dryers found")
                    st.dataframe(matched)
                    col = "equipment id" if "equipment id" in matched.columns else matched.columns[0]
                    sel = st.selectbox("Select dryer:", matched[col].astype(str).tolist(), key=f"seld_{batch_id}")
                    st.session_state.selections.append(
                        ([{"unit_op": batch_id, "operation": "drying", "material": "N/A",
                           "input_volume": 0, "actual_volume": vol, "accumulated_volume": vol}], sel))
                else:
                    st.warning("No matching dryers found.")

    if st.session_state.selections:
        buf = export_steps_to_excel(st.session_state.selections)
        st.download_button("Download Steps Summary", data=buf.getvalue(),
                           file_name="unit_op_steps.xlsx")


# ==========================================================================
#  TAB 2 — PFD Generator (full 5-step Python pipeline + formatter)
# ==========================================================================
def tab_pfd():
    st.header("PFD Generator — Report → Formatted Process Flow Diagram")
    st.caption("Python port of the 5 MATLAB steps + VBA formatter. "
               "Upload the optimization report, optionally a material sheet, "
               "then generate a formatted flowchart Excel.")

    # ---- Ingest report (Step 1 + 2) -------------------------------------
    col_a, col_b = st.columns(2)
    with col_a:
        report_file = st.file_uploader("Optimization report (PDF / DOCX)",
                                       type=["pdf", "docx"], key="pfd_report")
    with col_b:
        pasted = st.text_area("…or paste the numbered process steps here",
                              height=160, key="pfd_paste",
                              placeholder="1. Charge water lot-1 into the RBF\n"
                                          "2. Charge L-Tyrosine into the RBF\n"
                                          "Note: heterogeneous mass\n...")

    material_file = st.file_uploader(
        "Material sheet (optional .xlsx) — names in col B, Qty(Kg) in col E, Vol(L) in col F",
        type=["xlsx"], key="pfd_material")

    with st.expander("Advanced: condition keywords & column mapping"):
        kw_text = st.text_input(
            "Condition keywords (comma-separated)",
            value="If IPC, If complies, If does not comply, If IPC result complies")
        c1, c2, c3 = st.columns(3)
        name_col = c1.number_input("Material name column (1=A)", 1, 26, 2) - 1
        kg_col = c2.number_input("Qty Kg column (1=A)", 1, 26, 5) - 1
        l_col = c3.number_input("Vol L column (1=A)", 1, 26, 6) - 1

    if st.button("① Parse report", key="pfd_parse"):
        steps = []
        if report_file is not None:
            if report_file.name.lower().endswith(".pdf"):
                steps = pfd.extract_steps_from_pdf(report_file)
            else:
                steps = pfd.extract_steps_from_docx(report_file)
        elif pasted.strip():
            steps = pfd.extract_steps_from_text(pasted)

        if not steps:
            st.warning("No steps found. Upload a report or paste numbered steps.")
            return

        materials = None
        if material_file is not None:
            materials = pfd.load_material_sheet(
                material_file, name_col=int(name_col), kg_col=int(kg_col), l_col=int(l_col))

        keywords = [k.strip() for k in kw_text.split(",") if k.strip()]
        pfd.extract_conditions(steps, keywords)
        if materials:
            pfd.assign_materials(steps, materials)

        # store partially-processed steps for the volume stage
        st.session_state.pfd_steps = steps
        st.session_state.pfd_materials_loaded = bool(materials)
        st.success(f"Parsed {len(steps)} steps"
                   + (f", loaded {len(materials)} materials" if materials else ""))

    # ---- Step 5 : volume ledger (with 'separate' prompts) ----------------
    if "pfd_steps" in st.session_state:
        steps = st.session_state.pfd_steps
        st.markdown("### Extracted steps")
        st.dataframe(pfd.steps_to_dataframe(steps), use_container_width=True, height=300)

        sep_steps = pfd.separate_steps(steps)
        overrides = {}
        if sep_steps:
            st.markdown("### 'Separate' steps — enter Min Vol (L) for each reset point")
            for s in sep_steps:
                overrides[s.number] = st.number_input(
                    f"Step {s.number}: {s.text[:70]}",
                    min_value=0.0, value=0.0, key=f"sep_{s.number}")

        renderer = st.radio("Flowchart renderer",
                            ["openpyxl (boxes + fills + arrows)",
                             "xlsxwriter (merged flowchart boxes)"],
                            horizontal=True, key="pfd_renderer")

        if st.button("② Compute volumes & generate PFD Excel", key="pfd_generate"):
            pfd.compute_volumes(steps, separate_overrides=overrides)
            st.dataframe(pfd.steps_to_dataframe(steps), use_container_width=True, height=300)

            if renderer.startswith("openpyxl"):
                buf = io.BytesIO()
                pfd.write_formatted_workbook(steps, buf)
                buf.seek(0)
                data = buf.getvalue()
            else:
                data = pfd.write_flowchart_xlsxwriter(steps).getvalue()

            st.success("PFD generated.")
            st.download_button("Download PFD Excel", data=data,
                               file_name="PFD_output.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ==========================================================================
#  TAB 3 — Scale-up + Mass Balance (CTGR3-style PFD)
# ==========================================================================
def tab_scaleup():
    st.header("Scale-up & Mass Balance — CTGR3-style PFD")
    st.caption("Enter the lab recipe + manual scale-up ratio and yield. "
               "Materials are simple-scaled (plant = lab × ratio); flag "
               "stoichiometric/solution materials as manual and type their plant qty. "
               "Output matches the CTGR3.xlsx layout (top block + table + flowchart).")

    n_stages = st.number_input("Number of stages", 1, 5, 1, key="su_nstages")

    stages = []
    for si in range(1, int(n_stages) + 1):
        with st.expander(f"Stage {si}", expanded=(si == 1)):
            c1, c2, c3 = st.columns(3)
            name = c1.text_input("Sheet name", value=f"Stage{si}", key=f"su_name_{si}")
            project = c2.text_input("Project", value="Ajinomoto", key=f"su_proj_{si}")
            ratio = c3.number_input("Scale-up ratio (manual)", min_value=0.0,
                                    value=410.0, key=f"su_ratio_{si}")
            c4, c5, c6, c7 = st.columns(4)
            lab_input = c4.number_input("Lab input (g)", min_value=0.0, value=100.0, key=f"su_labin_{si}")
            in_mw = c5.number_input("Input MW", min_value=0.0, value=181.19, key=f"su_inmw_{si}")
            prod_mw = c6.number_input("Product MW", min_value=0.0, value=257.67, key=f"su_pmw_{si}")
            yld = c7.number_input("Yield (fraction)", min_value=0.0, max_value=1.0,
                                  value=0.8, key=f"su_yield_{si}")
            batch_override = st.number_input(
                "Plant batch size (kg) — leave 0 to auto (lab×ratio or previous output)",
                min_value=0.0, value=0.0, key=f"su_batch_{si}")

            st.markdown("**Raw materials** — edit the grid "
                        "(mark `manual` for stoichiometric/solution rows and fill `manual_kg`/`manual_L`)")
            default_rows = pd.DataFrame([
                {"sr_no": 1, "name": "L-Tyrosine", "lab_g": 100.0, "lab_ml": 0.0,
                 "mw": 181.19, "density": 1.6, "remarks": "SM, 1.0 eq",
                 "manual": False, "manual_kg": 0.0, "manual_L": 0.0},
                {"sr_no": 2, "name": "Water Lot-1", "lab_g": 0.0, "lab_ml": 300.0,
                 "mw": 18.0, "density": 0.0, "remarks": "3 vol",
                 "manual": False, "manual_kg": 0.0, "manual_L": 0.0},
                {"sr_no": 3, "name": "NaOH Lot-1", "lab_g": 0.0, "lab_ml": 0.0,
                 "mw": 40.0, "density": 2.1, "remarks": "2.0 eq (stoich)",
                 "manual": True, "manual_kg": 18.1, "manual_L": 0.0},
            ])
            grid = st.data_editor(default_rows, num_rows="dynamic",
                                  use_container_width=True, key=f"su_grid_{si}")

            report_txt = st.text_area("Process steps for the flowchart (numbered)",
                                      height=140, key=f"su_steps_{si}",
                                      placeholder="1. Charge Water Lot-1 into the reactor\n"
                                                  "2. Charge L-Tyrosine into the reactor\n...")

            # build stage object
            mats = []
            for _, r in grid.iterrows():
                if not str(r["name"]).strip():
                    continue
                mats.append(scaleup.RawMaterial(
                    sr_no=int(r["sr_no"]) if pd.notna(r["sr_no"]) else 0,
                    name=str(r["name"]).strip(),
                    lab_qty_g=float(r["lab_g"]) or None,
                    lab_vol_ml=float(r["lab_ml"]) or None,
                    mw=float(r["mw"]) or None,
                    density=float(r["density"]) or None,
                    remarks=str(r["remarks"]),
                    manual=bool(r["manual"]),
                    manual_plant_kg=float(r["manual_kg"]) or None,
                    manual_plant_l=float(r["manual_L"]) or None,
                ))
            stage = scaleup.StageCalc(
                name=name, project=project, lab_input_g=lab_input, input_mw=in_mw,
                product_mw=prod_mw, scale_up_ratio=ratio, yield_frac=yld,
                plant_batch_kg=(batch_override or None), materials=mats,
                steps=pfd.extract_steps_from_text(report_txt) if report_txt.strip() else [],
            )
            stages.append(stage)

    if st.button("Compute & generate CTGR3-style workbook", key="su_generate"):
        scaleup.chain_stages(stages)

        # attach materials + volumes to each stage's flowchart steps
        for stage in stages:
            if not stage.steps:
                continue
            pfd.extract_conditions(stage.steps)
            mat_rows = [[m.sr_no, m.name, None, m.plant_kg, m.plant_l] for m in stage.materials]
            mats = pfd.load_material_sheet(pd.DataFrame(mat_rows), name_col=1, kg_col=3, l_col=4)
            pfd.assign_materials(stage.steps, mats)
            overrides = {s.number: 0.0 for s in pfd.separate_steps(stage.steps)}
            pfd.compute_volumes(stage.steps, separate_overrides=overrides)

        # show summary
        for stage in stages:
            st.subheader(stage.name)
            st.write(f"Batch size: **{stage.plant_batch_kg:.3f} kg** · "
                     f"Input moles: **{stage.input_moles:.5f}** · "
                     f"Output: **{stage.output_kg:.3f} kg**")
            rows = [{"Sr": m.sr_no, "Material": m.name,
                     "Lab (g)": m.lab_qty_g, "Lab (ml)": m.lab_vol_ml,
                     "Plant (Kg)": round(m.plant_kg, 3) if m.plant_kg else None,
                     "Plant (L)": round(m.plant_l, 3) if m.plant_l else None,
                     "Moles": round(m.moles, 4) if m.moles else None,
                     "W/w": round(m.w_w, 4) if m.w_w else None,
                     "V/W": round(m.v_w, 3) if m.v_w else None,
                     "Remarks": m.remarks} for m in stage.materials]
            st.dataframe(pd.DataFrame(rows), use_container_width=True)

        buf = io.BytesIO()
        scaleup.build_workbook(stages, buf)
        buf.seek(0)
        st.download_button("Download CTGR3-style PFD Excel", data=buf.getvalue(),
                           file_name="PFD_scaleup.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ==========================================================================
#  TAB 0 — One-click: upload PR&D report -> full PFD
# ==========================================================================
def tab_oneclick():
    st.header("Generate PFD from PR&D Report")
    st.caption("Upload the PR&D report PDF. The app auto-extracts the process steps, "
               "raw-materials table, molecular weights and yield, applies your scale-up "
               "ratio, and generates the CTGR3-style PFD workbook.")

    report = st.file_uploader("PR&D report (PDF)", type=["pdf"], key="oc_report")

    c1, c2 = st.columns(2)
    ratio = c1.number_input("Scale-up ratio (plant ÷ lab)", min_value=0.0,
                            value=410.0, key="oc_ratio",
                            help="The one plant decision not contained in the report.")
    project = c2.text_input("Project name", value="", key="oc_project")

    if report is not None and st.button("Generate PFD", key="oc_go", type="primary"):
        with st.spinner("Parsing report…"):
            parsed = rparse.parse_report(report)

        if not parsed:
            st.error("Couldn't read the report. Is it a text-based PDF? "
                     "If it's a scan, use the 'Scale-up & Mass Balance' tab to enter data manually.")
            return

        stage_calcs = [rparse.to_stage_calc(ps, scale_up_ratio=ratio, project=project)
                       for ps in parsed]
        scaleup.chain_stages(stage_calcs)

        # attach materials + volume ledger to each flowchart
        for stage in stage_calcs:
            if not stage.steps:
                continue
            pfd.extract_conditions(stage.steps)
            mrows = [[m.sr_no, m.name, None, m.plant_kg, m.plant_l] for m in stage.materials]
            mats = pfd.load_material_sheet(pd.DataFrame(mrows), name_col=1, kg_col=3, l_col=4)
            pfd.assign_materials(stage.steps, mats)
            pfd.compute_volumes(
                stage.steps,
                separate_overrides={s.number: 0.0 for s in pfd.separate_steps(stage.steps)})

        st.success(f"Extracted {len(stage_calcs)} stage(s).")
        for ps, sc in zip(parsed, stage_calcs):
            with st.expander(f"{sc.name} — review extracted data", expanded=True):
                warn = []
                if ps.input_mw is None or ps.product_mw is None:
                    warn.append("molecular weights not found in scheme")
                if ps.yield_frac is None:
                    warn.append("yield not found (defaulted to 0.8)")
                if not ps.materials:
                    warn.append("raw-materials table not detected")
                if warn:
                    st.warning("Check: " + "; ".join(warn))
                st.write(f"Batch **{sc.plant_batch_kg:.2f} kg** · "
                         f"input moles **{sc.input_moles:.5f}** · "
                         f"yield **{sc.yield_frac:.3f}** · output **{sc.output_kg:.2f} kg**")
                rows = [{"Sr": m.sr_no, "Material": m.name,
                         "Lab g": m.lab_qty_g, "Lab mL": m.lab_vol_ml,
                         "Plant Kg": round(m.plant_kg, 3) if m.plant_kg else None,
                         "Plant L": round(m.plant_l, 3) if m.plant_l else None,
                         "MW": m.mw, "Density": m.density,
                         "Moles": round(m.moles, 4) if m.moles else None,
                         "V/W": round(m.v_w, 3) if m.v_w else None,
                         "Remarks": m.remarks} for m in sc.materials]
                if rows:
                    st.dataframe(pd.DataFrame(rows), use_container_width=True)

        buf = io.BytesIO()
        # formula-linked workbook: the whole sheet is driven by the ratio cell,
        # so a plant engineer can change the ratio in Excel and everything recomputes.
        fwriter.build_formula_workbook(stage_calcs, [ratio] * len(stage_calcs), buf)
        buf.seek(0)
        st.download_button("⬇ Download PFD Excel (formula-linked)", data=buf.getvalue(),
                           file_name="PFD.xlsx", type="primary",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("The workbook is fully formula-linked: change the Scale-Up ratio "
                   "cell (I12) in Excel and the batch size, plant quantities and volume "
                   "ledger all recalculate. Fill the Equipment ID + Capacity cells to get "
                   "occupancy %; the Observations column is left blank for manual entry.")


# ==========================================================================
def main():
    tab0, tab1, tab2, tab3 = st.tabs(
        ["① Report → PFD", "Equipment Selection", "PFD Generator (steps only)",
         "Scale-up (manual entry)"])
    with tab0:
        tab_oneclick()
    with tab1:
        tab_equipment()
    with tab2:
        tab_pfd()
    with tab3:
        tab_scaleup()


if __name__ == "__main__":
    main()
