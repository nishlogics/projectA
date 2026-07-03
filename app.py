"""
Process Flow Diagram (PFD) generator — Streamlit app.

Upload a PR&D report (PDF or Word) and it auto-detects the raw-material table,
molecular weights and yield, computes the mass balance, and generates a minimalist
PFD workbook (real rectangles for operations, diamonds for decisions, arrows).

A second tab provides equipment selection (reactor / filter / dryer).
"""

import io

import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

import pfd_pipeline as pfd
import pfd_scaleup as scaleup
import report_parser as rparse
import pfd_shapes as shapes

st.set_page_config(page_title="PFD Generator", layout="wide")
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


# ==========================================================================
#  PFD tab — upload report -> generate PFD
# ==========================================================================
def tab_pfd():
    st.header("Generate PFD from report")
    st.write("Upload a PR&D report (PDF or Word). The app reads the raw-material "
             "table, molecular weights and yield, then builds the flow diagram.")

    report = st.file_uploader("Report (PDF or Word)", type=["pdf", "docx"])

    col1, col2 = st.columns(2)
    batch_size = col1.number_input(
        "Plant batch size in Kg (optional — leave 0 to use the report quantities as-is)",
        min_value=0.0, value=0.0,
        help="If the report is a lab recipe and you are scaling to a plant batch, "
             "enter the plant batch size of the starting material. Otherwise the "
             "quantities in the report are used directly.")
    project = col2.text_input("Project name (optional)", value="")

    if report is None:
        return

    if st.button("Generate PFD", type="primary"):
        with st.spinner("Reading report…"):
            if report.name.lower().endswith(".docx"):
                parsed = rparse.parse_docx(report)
            else:
                parsed = rparse.parse_report(report)

        if not parsed:
            st.error("Could not read a raw-material table from this report. "
                     "If it is a scanned image PDF, the text can't be extracted.")
            return

        stages = []
        for ps in parsed:
            # scale-up ratio derived from the batch size if given, else 1:1
            sm_lab = ps.lab_input_g or (ps.materials[0].lab_qty_g if ps.materials else None)
            ratio = 1.0
            if batch_size and sm_lab:
                ratio = batch_size * 1000.0 / sm_lab   # kg -> g then /lab g
            sc = rparse.to_stage_calc(ps, scale_up_ratio=ratio, project=project)
            stages.append(sc)

        scaleup.chain_stages(stages)

        for stage in stages:
            if not stage.steps:
                continue
            pfd.extract_conditions(stage.steps)
            mrows = [[m.sr_no, m.name, None, m.plant_kg, m.plant_l] for m in stage.materials]
            if mrows:
                mats = pfd.load_material_sheet(pd.DataFrame(mrows), name_col=1, kg_col=3, l_col=4)
                pfd.assign_materials(stage.steps, mats)
            pfd.compute_volumes(
                stage.steps,
                separate_overrides={s.number: 0.0 for s in pfd.separate_steps(stage.steps)})

        # show what was read, so the user can sanity-check
        for ps, sc in zip(parsed, stages):
            with st.expander(f"{sc.name} — data read from report", expanded=True):
                st.write(f"Batch **{sc.plant_batch_kg:g} Kg** · "
                         f"moles **{sc.input_moles:.4g}** · "
                         f"yield **{sc.yield_frac:g}** · output **{sc.output_kg:.4g} Kg**")
                rows = [{"S.No": m.sr_no, "Material": m.name,
                         "Qty": m.lab_qty_g or m.lab_vol_ml,
                         "Unit": "g" if m.lab_qty_g else ("mL" if m.lab_vol_ml else ""),
                         "MW": m.mw, "Mole ratio": m.remarks,
                         "Plant Kg": round(m.plant_kg, 3) if m.plant_kg else None,
                         "Plant L": round(m.plant_l, 3) if m.plant_l else None}
                        for m in sc.materials]
                if rows:
                    st.dataframe(pd.DataFrame(rows), use_container_width=True)
                if not sc.materials:
                    st.warning("No raw-material table detected for this stage.")

        data = shapes.build(stages)
        st.success("PFD generated.")
        st.download_button("Download PFD (Excel)", data=data, file_name="PFD.xlsx",
                           type="primary",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def main():
    tab1, tab2 = st.tabs(["PFD Generator", "Equipment Selection"])
    with tab1:
        tab_pfd()
    with tab2:
        tab_equipment()


if __name__ == "__main__":
    main()
