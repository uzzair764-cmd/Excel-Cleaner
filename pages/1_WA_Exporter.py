import os
import re
import tempfile
import streamlit as st

from processors.wa_processor import run_export


st.set_page_config(
    page_title="WA Exporter",
    page_icon="whatsapp.png",
    layout="wide"
)

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1200px;
    }

    div[data-testid="stSidebar"] {
        background-color: #fafafa;
    }

    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        margin: 0;
    }

    .sub-title {
        color: #666;
        margin-bottom: 2rem;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# RESET HANDLER — MUST RUN BEFORE WIDGETS ARE CREATED
# ============================================================

if st.session_state.get("RESET_FILTERS_NOW", False):
    keys_to_delete = [
        "kaum_filter",
        "custom_kaum",
        "sikap_filter",
        "custom_sikap",
        "party_filter_selected",
        "print_full_summary",
        "age_groups",
        "zip_bytes",
        "summary_text",
        "final_rows",
        "files_created",
        "RESET_FILTERS_NOW",
    ]

    for key in keys_to_delete:
        if key in st.session_state:
            del st.session_state[key]

    st.session_state["age_groups"] = []


# ============================================================
# HELPERS
# ============================================================

def safe_filename(name):
    name = os.path.basename(str(name))
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    return name


def csv_to_list(text):
    text = str(text).strip()

    if text == "" or text == "NONE" or text == "NO FILTER":
        return []

    return [x.strip().upper() for x in text.split(",") if x.strip() != ""]


def normalize_bucket(x):
    x = str(x).strip().upper()
    x = x.replace("_", " ")
    x = x.replace("LAIN LAIN", "LAIN-LAIN")
    x = x.replace("LAINLAIN", "LAIN-LAIN")

    if x not in ["MELAYU", "CINA", "INDIA", "LAIN-LAIN"]:
        return "LAIN-LAIN"

    return x


def normalize_sikap_config(x):
    x = str(x).strip().upper()
    x = x.replace("_", " ")

    if "KELABU" in x:
        return "KELABU"

    if "HITAM" in x:
        return "HITAM"

    if "PUTIH" in x:
        return "PUTIH"

    return x


def get_structure_options(input_level):
    if input_level == "DUN":
        return {
            "DUN > DM.xlsx": "DUN_DM_FILE",
            "DUN > DM > KAUM": "DUN_DM_KAUM",
            "DUN > DUN.xlsx": "DUN_FILE",
            "DUN > KAUM": "DUN_KAUM",
            "DUN > AGE RANGE": "DUN_AGE",
            "DUN > DM > ML/MP/...": "DUN_DM_CODE",
            "COMBINE ALL INPUT FILES > ONE XLSX": "COMBINE_ALL",
        }

    return {
        "PARLIMEN > DUN > DM.xlsx": "PARLIMEN_DUN_DM_FILE",
        "PARLIMEN > DUN > DM > KAUM": "PARLIMEN_DUN_DM_KAUM",
        "PARLIMEN > DUN.xlsx": "PARLIMEN_DUN_FILE",
        "PARLIMEN > DUN > KAUM": "PARLIMEN_DUN_KAUM",
        "PARLIMEN > DUN > AGE RANGE": "PARLIMEN_DUN_AGE",
        "PARLIMEN > DUN > DM > ML/MP/...": "PARLIMEN_DUN_DM_CODE",
        "PARLIMEN > PARLIMEN.xlsx": "PARLIMEN_FILE",
        "COMBINE ALL INPUT FILES > ONE XLSX": "COMBINE_ALL",
    }


def map_structure_to_config(structure_code):
    force_split_by_kaum = False
    force_split_by_code = False

    if structure_code == "COMBINE_ALL":
        group_levels = []
        last_group_as_folder = False

    elif structure_code == "DUN_DM_FILE":
        group_levels = ["DUN", "DM"]
        last_group_as_folder = False

    elif structure_code == "DUN_DM_KAUM":
        group_levels = ["DUN", "DM"]
        last_group_as_folder = True
        force_split_by_kaum = True

    elif structure_code == "DUN_FILE":
        group_levels = ["DUN"]
        last_group_as_folder = True

    elif structure_code == "DUN_KAUM":
        group_levels = ["DUN"]
        last_group_as_folder = True
        force_split_by_kaum = True

    elif structure_code == "DUN_AGE":
        group_levels = ["DUN", "AGE"]
        last_group_as_folder = False

    elif structure_code == "DUN_DM_CODE":
        group_levels = ["DUN", "DM"]
        last_group_as_folder = True
        force_split_by_code = True

    elif structure_code == "PARLIMEN_DUN_DM_FILE":
        group_levels = ["PARLIMEN", "DUN", "DM"]
        last_group_as_folder = False

    elif structure_code == "PARLIMEN_DUN_DM_KAUM":
        group_levels = ["PARLIMEN", "DUN", "DM"]
        last_group_as_folder = True
        force_split_by_kaum = True

    elif structure_code == "PARLIMEN_DUN_FILE":
        group_levels = ["PARLIMEN", "DUN"]
        last_group_as_folder = False

    elif structure_code == "PARLIMEN_DUN_KAUM":
        group_levels = ["PARLIMEN", "DUN"]
        last_group_as_folder = True
        force_split_by_kaum = True

    elif structure_code == "PARLIMEN_DUN_AGE":
        group_levels = ["PARLIMEN", "DUN", "AGE"]
        last_group_as_folder = False

    elif structure_code == "PARLIMEN_DUN_DM_CODE":
        group_levels = ["PARLIMEN", "DUN", "DM"]
        last_group_as_folder = True
        force_split_by_code = True

    elif structure_code == "PARLIMEN_FILE":
        group_levels = ["PARLIMEN"]
        last_group_as_folder = True

    else:
        group_levels = ["DUN", "DM"]
        last_group_as_folder = False

    return group_levels, last_group_as_folder, force_split_by_kaum, force_split_by_code


def build_keep_kaum(kaum_filter, custom_kaum):
    if kaum_filter == "ALL":
        return ["ALL"]

    if kaum_filter == "MCI":
        return ["MELAYU", "CINA", "INDIA"]

    if kaum_filter in ["MELAYU", "CINA", "INDIA", "LAIN-LAIN"]:
        return [kaum_filter]

    custom = csv_to_list(custom_kaum)
    custom = [normalize_bucket(x) for x in custom]
    custom = [x for x in custom if x in ["MELAYU", "CINA", "INDIA", "LAIN-LAIN"]]

    return custom if custom else ["ALL"]


def build_sikap_filter(sikap_filter, custom_sikap):
    if sikap_filter == "NONE":
        return []

    if sikap_filter == "HITAM":
        return ["HITAM"]

    if sikap_filter == "KELABU":
        return ["KELABU"]

    if sikap_filter == "PUTIH":
        return ["PUTIH"]

    if sikap_filter == "HITAM_KELABU":
        return ["HITAM", "KELABU"]

    if sikap_filter == "KELABU_PUTIH":
        return ["KELABU", "PUTIH"]

    return [normalize_sikap_config(x) for x in csv_to_list(custom_sikap)]


def parse_age_groups(age_group_rows):
    age_ranges = []

    for row in age_group_rows:
        label = str(row.get("label", "")).strip()
        min_age = str(row.get("min", "")).strip()
        max_age = str(row.get("max", "")).strip()

        if label == "" and min_age == "" and max_age == "":
            continue

        if min_age == "":
            continue

        try:
            min_age_int = int(min_age)
        except:
            continue

        if max_age == "":
            max_age_int = None
        else:
            try:
                max_age_int = int(max_age)
            except:
                continue

        if label == "":
            if max_age_int is None:
                label = f"{min_age_int}>"
                file_label = f"{min_age_int} KE ATAS"
            else:
                label = f"{min_age_int}-{max_age_int}"
                file_label = f"{min_age_int}-{max_age_int}"
        else:
            file_label = label

        age_ranges.append((label, file_label, min_age_int, max_age_int))

    return age_ranges


def get_preview_age_labels(age_ranges):
    if age_ranges:
        return [x[1] for x in age_ranges]

    return ["ADD AGE GROUP FIRST"]


def build_folder_preview(input_level, structure_code, age_ranges):
    age_labels = get_preview_age_labels(age_ranges)

    if structure_code == "COMBINE_ALL":
        return """voter_outputs/
└── OUTPUT.xlsx"""

    if input_level == "DUN":
        if structure_code == "DUN_DM_FILE":
            return """voter_outputs/
└── N.01 DUN_NAME/
    ├── DM01 NAMA_DM.xlsx
    ├── DM02 NAMA_DM.xlsx
    └── DM03 NAMA_DM.xlsx"""

        if structure_code == "DUN_DM_KAUM":
            return """voter_outputs/
└── N.01 DUN_NAME/
    └── DM01 NAMA_DM/
        ├── DM01 NAMA_DM MELAYU.xlsx
        ├── DM01 NAMA_DM CINA.xlsx
        ├── DM01 NAMA_DM INDIA.xlsx
        └── DM01 NAMA_DM LAIN-LAIN.xlsx"""

        if structure_code == "DUN_FILE":
            return """voter_outputs/
├── N.01 DUN_NAME.xlsx
├── N.02 DUN_NAME.xlsx
└── N.03 DUN_NAME.xlsx"""

        if structure_code == "DUN_KAUM":
            return """voter_outputs/
└── N.01 DUN_NAME/
    ├── N.01 DUN_NAME MELAYU.xlsx
    ├── N.01 DUN_NAME CINA.xlsx
    ├── N.01 DUN_NAME INDIA.xlsx
    └── N.01 DUN_NAME LAIN-LAIN.xlsx"""

        if structure_code == "DUN_AGE":
            lines = ["voter_outputs/", "└── N.01 DUN_NAME/"]
            for i, age in enumerate(age_labels):
                prefix = "    └──" if i == len(age_labels) - 1 else "    ├──"
                lines.append(f"{prefix} {age}.xlsx")
            return "\n".join(lines)

        if structure_code == "DUN_DM_CODE":
            return """voter_outputs/
└── N.01 DUN_NAME/
    └── DM01 NAMA_DM/
        ├── DM01 NAMA_DM ML.xlsx
        ├── DM01 NAMA_DM MP.xlsx
        ├── DM01 NAMA_DM CL.xlsx
        └── DM01 NAMA_DM CP.xlsx"""

    else:
        if structure_code == "PARLIMEN_DUN_DM_FILE":
            return """voter_outputs/
└── P001 PARLIMEN_NAME/
    └── N.01 DUN_NAME/
        ├── DM01 NAMA_DM.xlsx
        └── DM02 NAMA_DM.xlsx"""

        if structure_code == "PARLIMEN_DUN_DM_KAUM":
            return """voter_outputs/
└── P001 PARLIMEN_NAME/
    └── N.01 DUN_NAME/
        └── DM01 NAMA_DM/
            ├── DM01 NAMA_DM MELAYU.xlsx
            ├── DM01 NAMA_DM CINA.xlsx
            └── DM01 NAMA_DM LAIN-LAIN.xlsx"""

        if structure_code == "PARLIMEN_DUN_FILE":
            return """voter_outputs/
└── P001 PARLIMEN_NAME/
    ├── N.01 DUN_NAME.xlsx
    └── N.02 DUN_NAME.xlsx"""

        if structure_code == "PARLIMEN_DUN_KAUM":
            return """voter_outputs/
└── P001 PARLIMEN_NAME/
    └── N.01 DUN_NAME/
        ├── N.01 DUN_NAME MELAYU.xlsx
        ├── N.01 DUN_NAME CINA.xlsx
        └── N.01 DUN_NAME LAIN-LAIN.xlsx"""

        if structure_code == "PARLIMEN_DUN_AGE":
            lines = ["voter_outputs/", "└── P001 PARLIMEN_NAME/", "    └── N.01 DUN_NAME/"]
            for i, age in enumerate(age_labels):
                prefix = "        └──" if i == len(age_labels) - 1 else "        ├──"
                lines.append(f"{prefix} {age}.xlsx")
            return "\n".join(lines)

        if structure_code == "PARLIMEN_DUN_DM_CODE":
            return """voter_outputs/
└── P001 PARLIMEN_NAME/
    └── N.01 DUN_NAME/
        └── DM01 NAMA_DM/
            ├── DM01 NAMA_DM ML.xlsx
            ├── DM01 NAMA_DM MP.xlsx
            └── DM01 NAMA_DM CL.xlsx"""

        if structure_code == "PARLIMEN_FILE":
            return """voter_outputs/
├── P001 PARLIMEN_NAME.xlsx
├── P002 PARLIMEN_NAME.xlsx
└── P003 PARLIMEN_NAME.xlsx"""

    return "voter_outputs/"


def request_filter_reset():
    st.session_state["RESET_FILTERS_NOW"] = True
    st.rerun()


if "age_groups" not in st.session_state:
    st.session_state["age_groups"] = []


# ============================================================
# HEADER
# ============================================================

col_logo, col_title = st.columns([0.07, 0.93])

with col_logo:
    st.image("whatsapp.png", width=54)

with col_title:
    st.markdown('<div class="main-title">WA Exporter</div>', unsafe_allow_html=True)

st.markdown(
    '<div class="sub-title">Clean Excel files into WhatsApp import format with DUN / DM / KAUM / AGE filters.</div>',
    unsafe_allow_html=True
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("Configurations")

    input_level = st.selectbox(
        "Input Level",
        ["DUN", "PARLIMEN"],
        index=0,
        key="input_level"
    )

    structure_options = get_structure_options(input_level)

    if st.session_state.get("structure_label") not in structure_options:
        st.session_state["structure_label"] = list(structure_options.keys())[0]

    structure_label = st.selectbox(
        "Output Structure",
        list(structure_options.keys()),
        key="structure_label"
    )

    structure_code = structure_options[structure_label]

    st.divider()

    kaum_filter = st.selectbox(
        "Kaum Filter",
        ["ALL", "MCI", "MELAYU", "CINA", "INDIA", "LAIN-LAIN", "CUSTOM"],
        format_func=lambda x: {
            "ALL": "MELAYU, CINA, INDIA, LAIN-LAIN",
            "MCI": "MELAYU, CINA, INDIA",
            "MELAYU": "MELAYU",
            "CINA": "CINA",
            "INDIA": "INDIA",
            "LAIN-LAIN": "LAIN-LAIN",
            "CUSTOM": "CUSTOM",
        }[x],
        index=0,
        key="kaum_filter"
    )

    if kaum_filter == "CUSTOM":
        custom_kaum = st.text_input(
            "Custom Kaum",
            placeholder="Example: MELAYU,CINA,INDIA,LAIN-LAIN",
            key="custom_kaum"
        )
    else:
        custom_kaum = ""

    sikap_filter = st.selectbox(
        "Sikap Filter",
        ["NONE", "HITAM", "KELABU", "PUTIH", "HITAM_KELABU", "KELABU_PUTIH", "CUSTOM"],
        format_func=lambda x: {
            "NONE": "NO FILTER",
            "HITAM": "HITAM only",
            "KELABU": "KELABU only",
            "PUTIH": "PUTIH only",
            "HITAM_KELABU": "HITAM + KELABU",
            "KELABU_PUTIH": "KELABU + PUTIH",
            "CUSTOM": "Custom sikap filter",
        }[x],
        index=0,
        key="sikap_filter"
    )

    if sikap_filter == "CUSTOM":
        custom_sikap = st.text_input(
            "Custom Sikap",
            placeholder="Example: HITAM,KELABU,PUTIH",
            key="custom_sikap"
        )
    else:
        custom_sikap = ""

    party_filter_selected = st.selectbox(
        "Party / Parti Filter",
        ["NO FILTER", "PAS", "PKR", "PPBM", "UMNO"],
        index=0,
        key="party_filter_selected"
    )

    st.divider()

    st.write("Age Groups")

    if st.button("➕ Add age group", use_container_width=True):
        next_no = len(st.session_state["age_groups"]) + 1
        st.session_state["age_groups"].append({
            "id": next_no,
            "label": "",
            "min": "",
            "max": "",
        })
        st.rerun()

    updated_age_groups = []

    for i, row in enumerate(st.session_state["age_groups"]):
        with st.container(border=True):
            st.caption(f"Age Group {i + 1}")

            label = st.text_input(
                "Label",
                value=row.get("label", ""),
                placeholder="Example: 18-40",
                key=f"age_label_{row['id']}"
            )

            col_age_1, col_age_2 = st.columns(2)

            with col_age_1:
                min_age_input = st.text_input(
                    "Min",
                    value=row.get("min", ""),
                    placeholder="18",
                    key=f"age_min_{row['id']}"
                )

            with col_age_2:
                max_age_input = st.text_input(
                    "Max",
                    value=row.get("max", ""),
                    placeholder="Blank = no max",
                    key=f"age_max_{row['id']}"
                )

            remove_clicked = st.button(
                "Remove",
                key=f"remove_age_{row['id']}",
                use_container_width=True
            )

            if not remove_clicked:
                updated_age_groups.append({
                    "id": row["id"],
                    "label": label,
                    "min": min_age_input,
                    "max": max_age_input,
                })

    if updated_age_groups != st.session_state["age_groups"]:
        st.session_state["age_groups"] = updated_age_groups
        st.rerun()

    custom_age_ranges = parse_age_groups(st.session_state["age_groups"])

    st.divider()

    print_full_summary = st.checkbox(
        "Show full summary after processing",
        value=False,
        key="print_full_summary"
    )

    st.divider()

    st.button(
        "RESET FILTERS",
        use_container_width=True,
        on_click=request_filter_reset
    )


# ============================================================
# FOLDER PREVIEW
# ============================================================

st.markdown("### Folder Structure Preview")

preview_text = build_folder_preview(
    input_level=input_level,
    structure_code=structure_code,
    age_ranges=custom_age_ranges
)

st.code(preview_text, language="text")


# ============================================================
# FILE UPLOAD + RUN
# ============================================================

uploaded_files = st.file_uploader(
    "Upload Excel file(s)",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

run_button = st.button("Run Export", type="primary", use_container_width=True)


if run_button:
    if not uploaded_files:
        st.error("Upload at least one Excel file first.")
        st.stop()

    group_levels, last_group_as_folder, force_split_by_kaum, force_split_by_code = map_structure_to_config(structure_code)

    split_by_kaum = force_split_by_kaum
    split_by_gender_code = force_split_by_code

    keep_kaum = build_keep_kaum(kaum_filter, custom_kaum)
    sikap_filter_list = build_sikap_filter(sikap_filter, custom_sikap)

    if party_filter_selected == "NO FILTER":
        party_filter = []
    else:
        party_filter = [party_filter_selected]

    if custom_age_ranges:
        age_filter = "CUSTOM_AGE_GROUPS"
    else:
        age_filter = None

    with tempfile.TemporaryDirectory() as temp_dir:
        input_dir = os.path.join(temp_dir, "input")
        output_dir = os.path.join(temp_dir, "voter_outputs")

        os.makedirs(input_dir, exist_ok=True)

        file_paths = []

        for i, uploaded_file in enumerate(uploaded_files, start=1):
            clean_name = safe_filename(uploaded_file.name)
            file_path = os.path.join(input_dir, f"{i}_{clean_name}")

            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            file_paths.append(file_path)

        config = {
            "output_dir": output_dir,
            "zip_name": "voter_outputs",
            "input_level": input_level,
            "group_levels": group_levels,
            "last_group_as_folder": last_group_as_folder,
            "split_by_kaum": split_by_kaum,
            "split_by_gender_code": split_by_gender_code,
            "keep_kaum": keep_kaum,
            "sikap_filter": sikap_filter_list,
            "party_filter": party_filter,
            "age_filter": age_filter,
            "custom_age_ranges": custom_age_ranges,
            "dedup_by_phone": True,
            "dedup_by_nokp": False,
            "read_all_sheets": False,
            "create_empty_files": False,
        }

        progress_bar = st.progress(0)
        status_box = st.empty()

        def progress_callback(message, progress_value=None):
            status_box.info(message)

            if progress_value is not None:
                progress_bar.progress(min(max(int(progress_value), 0), 100))

        try:
            with st.spinner("Processing files..."):
                result = run_export(
                    file_paths=file_paths,
                    config=config,
                    progress_callback=progress_callback
                )

            progress_bar.progress(100)

            with open(result["zip_path"], "rb") as f:
                zip_bytes = f.read()

            st.session_state["zip_bytes"] = zip_bytes
            st.session_state["summary_text"] = result["summary_text"]
            st.session_state["final_rows"] = result["final_rows"]
            st.session_state["files_created"] = result["files_created"]
            st.session_state["zip_file_name"] = os.path.basename(result["zip_path"])

            status_box.empty()
            st.success("Done.")

        except Exception as e:
            status_box.empty()
            st.error(str(e))
            st.stop()


# ============================================================
# RESULT DISPLAY
# ============================================================

if "zip_bytes" in st.session_state:
    col1, col2 = st.columns(2)

    with col1:
        st.metric("Final rows", f"{st.session_state['final_rows']:,}")

    with col2:
        st.metric("Excel files created", f"{st.session_state['files_created']:,}")

    st.download_button(
        label="Download ZIP",
        data=st.session_state["zip_bytes"],
        file_name=st.session_state.get("zip_file_name", "voter_outputs.zip"),
        mime="application/zip",
        type="primary",
        use_container_width=True
    )

    if print_full_summary:
        with st.expander("View SUMMARY.txt", expanded=False):
            st.text(st.session_state["summary_text"])
