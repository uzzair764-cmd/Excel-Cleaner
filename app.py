import os
import re
import tempfile
import streamlit as st

from processor import run_export


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
        }

    return {
        "PARLIMEN > DUN > DM.xlsx": "PARLIMEN_DUN_DM_FILE",
        "PARLIMEN > DUN > DM > KAUM": "PARLIMEN_DUN_DM_KAUM",
        "PARLIMEN > DUN.xlsx": "PARLIMEN_DUN_FILE",
        "PARLIMEN > DUN > KAUM": "PARLIMEN_DUN_KAUM",
        "PARLIMEN > DUN > AGE RANGE": "PARLIMEN_DUN_AGE",
        "PARLIMEN > DUN > DM > ML/MP/...": "PARLIMEN_DUN_DM_CODE",
        "PARLIMEN > PARLIMEN.xlsx": "PARLIMEN_FILE",
    }


def map_structure_to_config(structure_code):
    force_split_by_kaum = False
    force_split_by_code = False

    if structure_code == "DUN_DM_FILE":
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


if "age_groups" not in st.session_state:
    st.session_state["age_groups"] = []


col_logo, col_title = st.columns([0.07, 0.93])

with col_logo:
    st.image("whatsapp.png", width=54)

with col_title:
    st.markdown('<div class="main-title">WA Exporter</div>', unsafe_allow_html=True)

st.markdown(
    '<div class="sub-title">Clean Excel files into WhatsApp import format with DUN / DM / KAUM / AGE filters.</div>',
    unsafe_allow_html=True
)


with st.sidebar:
    st.header("Settings")

    input_level = st.selectbox(
        "Input Level",
        ["DUN", "PARLIMEN"],
        index=0
    )

    structure_options = get_structure_options(input_level)
    structure_label = st.selectbox(
        "Output Structure",
        list(structure_options.keys()),
        index=0
    )
    structure_code = structure_options[structure_label]

    st.divider()

    kaum_filter = st.selectbox(
        "Kaum Filter",
        [
            "ALL",
            "MCI",
            "MELAYU",
            "CINA",
            "INDIA",
            "LAIN-LAIN",
            "CUSTOM",
        ],
         format_func=lambda x: {
            "ALL": "MELAYU, CINA, INDIA, LAIN-LAIN",
            "MCI": "MELAYU, CINA, INDIA",
            "MELAYU": "MELAYU",
            "CINA": "CINA",
            "INDIA": "INDIA",
            "LAIN-LAIN": "LAIN-LAIN",
            "CUSTOM": "CUSTOM",
        }[x],
        index=0
    )

    if kaum_filter == "CUSTOM":
        custom_kaum = st.text_input(
            "Custom Kaum",
            placeholder="Example: MELAYU,CINA,INDIA,LAIN-LAIN"
        )
    else:
        custom_kaum = ""

    sikap_filter = st.selectbox(
        "Sikap Filter",
        [
            "NONE",
            "HITAM",
            "KELABU",
            "PUTIH",
            "HITAM_KELABU",
            "KELABU_PUTIH",
            "CUSTOM",
        ],
        format_func=lambda x: {
            "NONE": "NO FILTER",
            "HITAM": "HITAM only",
            "KELABU": "KELABU only",
            "PUTIH": "PUTIH only",
            "HITAM_KELABU": "HITAM + KELABU",
            "KELABU_PUTIH": "KELABU + PUTIH",
            "CUSTOM": "Custom sikap filter",
        }[x],
        index=0
    )

    if sikap_filter == "CUSTOM":
        custom_sikap = st.text_input(
            "Custom Sikap",
            placeholder="Example: HITAM,KELABU,PUTIH"
        )
    else:
        custom_sikap = ""

    party_filter_selected = st.selectbox(
        "Party / Parti Filter",
        ["NO FILTER", "PAS", "PKR", "PPBM", "UMNO"],
        index=0
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

    print_full_summary = st.checkbox("Show full summary after processing", value=False)


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

        status_box = st.empty()

        def progress_callback(message):
            status_box.info(message)

        try:
            with st.spinner("Processing files..."):
                result = run_export(
                    file_paths=file_paths,
                    config=config,
                    progress_callback=progress_callback
                )

            with open(result["zip_path"], "rb") as f:
                zip_bytes = f.read()

            st.session_state["zip_bytes"] = zip_bytes
            st.session_state["summary_text"] = result["summary_text"]
            st.session_state["final_rows"] = result["final_rows"]
            st.session_state["files_created"] = result["files_created"]

            status_box.empty()
            st.success("Done.")

        except Exception as e:
            status_box.empty()
            st.error(str(e))
            st.stop()


if "zip_bytes" in st.session_state:
    col1, col2 = st.columns(2)

    with col1:
        st.metric("Final rows", f"{st.session_state['final_rows']:,}")

    with col2:
        st.metric("Excel files created", f"{st.session_state['files_created']:,}")

    st.download_button(
        label="Download ZIP",
        data=st.session_state["zip_bytes"],
        file_name="voter_outputs.zip",
        mime="application/zip",
        type="primary",
        use_container_width=True
    )

    if print_full_summary:
        with st.expander("View SUMMARY.txt", expanded=False):
            st.text(st.session_state["summary_text"])
