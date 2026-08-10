import io
import re

import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo


# ============================================================
# CONSTANTS
# ============================================================

HEADERS = [
    "KOD DM",
    "NAMA DM",
    "JUMLAH",

    "LELAKI",
    "LELAKI (%)",
    "PEREMPUAN",
    "PEREMPUAN (%)",

    "MELAYU",
    "MELAYU (%)",
    "CINA",
    "CINA (%)",
    "INDIA",
    "INDIA (%)",
    "LAIN-LAIN",
    "LAIN-LAIN (%)",

    "18-21",
    "18-21 (%)",
    "22-30",
    "22-30 (%)",
    "31-40",
    "31-40 (%)",
    "41-50",
    "41-50 (%)",
    "51-60",
    "51-60 (%)",
    "61+",
    "61+ (%)",

    "UMNO",
    "UMNO (%)",
    "PKR",
    "PKR (%)",
    "PAS",
    "PAS (%)",
    "PPBM",
    "PPBM (%)",

    "PUTIH",
    "PUTIH (%)",
    "KELABU",
    "KELABU (%)",
    "HITAM",
    "HITAM (%)",

    "PENGUNDI AWAL",
    "PENGUNDI AWAL (%)",

    "POLIS",
    "POLIS (%)",

    "PASANGAN POLIS",
    "PASANGAN POLIS (%)",

    "ASKAR",
    "ASKAR (%)",

    "PASANGAN ASKAR",
    "PASANGAN ASKAR (%)"
]


MAIN_RACES = [
    "MELAYU",
    "CINA",
    "INDIA",
    "LAIN-LAIN"
]

DEFAULT_AGE_GROUPS = [
    "18-21",
    "22-30",
    "31-40",
    "41-50",
    "51-60",
    "61+"
]

PARTY_COLS = [
    "UMNO",
    "PKR",
    "PAS",
    "PPBM"
]

SIKAP_COLS = [
    "PUTIH",
    "KELABU",
    "HITAM"
]


# ============================================================
# COLUMN FINDER
# ============================================================

def get_col(df, possible_names):
    """
    Find a dataframe column using case-insensitive matching.
    """

    col_map = {
        str(c).lower().strip(): c
        for c in df.columns
    }

    for name in possible_names:

        key = str(name).lower().strip()

        if key in col_map:
            return col_map[key]

    return None


# ============================================================
# CLEANING FUNCTIONS
# ============================================================

def clean_service_no(value):
    """
    Clean NoPerkhidmatan / service number.
    """

    if pd.isna(value):
        return ""

    n = str(value).strip().upper()

    if n in {
        "",
        "NAN",
        "NONE",
        "NULL"
    }:
        return ""

    return n


def clean_filename(value):
    """
    Make a safe filename.
    """

    name = str(value).strip().upper()

    name = re.sub(
        r'[\\/:*?"<>|]',
        " ",
        name
    )

    name = " ".join(name.split())

    return name if name else "OUTPUT"


def clean_sheet_name(value, existing_names):
    """
    Create a valid and unique Excel worksheet name.
    """

    name = str(value).strip().upper()

    # Excel-invalid worksheet characters
    name = re.sub(
        r'[\\/?*[\]:]',
        " ",
        name
    )

    name = " ".join(name.split())

    if not name:
        name = "DUN"

    name = name[:31]

    base = name
    counter = 2

    while name in existing_names:

        suffix = f" ({counter})"

        name = (
            base[:31 - len(suffix)]
            + suffix
        )

        counter += 1

    existing_names.add(name)

    return name


# ============================================================
# DUN FUNCTIONS
# ============================================================

def kod_dun_digits(kod_dun):
    """
    Extract only digits from KOD DUN.

    Handles values such as:
    14
    14.0
    N14
    N.14
    """

    if pd.isna(kod_dun):
        return ""

    kod_str = str(kod_dun).strip()

    # Handle Excel float representation
    kod_str = kod_str.split(".")[0]

    return re.sub(
        r"\D",
        "",
        kod_str
    )


def kod_dun_sort_key(kod_dun):
    digits = kod_dun_digits(kod_dun)

    if not digits:
        return -1

    return int(digits)


def format_sheet_label(kod_dun, nama_dun):
    """
    Example:
    kod_dun = 4
    nama_dun = KLAWANG

    Result:
    N.04 KLAWANG
    """

    digits = kod_dun_digits(kod_dun)

    last2 = (
        digits[-2:].zfill(2)
        if digits
        else "00"
    )

    return f"N.{last2} {nama_dun}"


# ============================================================
# KOD DM
# ============================================================

def format_kod_dm(value):
    """
    Convert a 7-digit KOD DM into:

    XXX/XX/XX

    Example:
    0400101
    ->
    040/01/01
    """

    if pd.isna(value):
        return ""

    kod = str(value).strip()

    if kod in {
        "",
        "None",
        "none",
        "nan",
        "NaN"
    }:
        return ""

    kod = kod.split(".")[0]

    kod = kod.zfill(7)

    return (
        f"{kod[:3]}/"
        f"{kod[3:5]}/"
        f"{kod[5:]}"
    )


# ============================================================
# RACE
# ============================================================

def normalise_race(value):

    if pd.isna(value):
        return "LAIN-LAIN"

    r = str(value).strip().upper()

    if r in {
        "MELAYU",
        "CINA",
        "INDIA"
    }:
        return r

    return "LAIN-LAIN"


# ============================================================
# SIKAP
# ============================================================

def normalise_sikap(value):

    if pd.isna(value):
        return ""

    s = str(value).strip().upper()

    if s in {
        "KELABU-LAMA",
        "KELABU-BARU"
    }:
        return "KELABU"

    if s in {
        "PUTIH",
        "KELABU",
        "HITAM"
    }:
        return s

    return ""


# ============================================================
# PENGUNDI AWAL
# ============================================================

def classify_awal(value):

    n = clean_service_no(value)

    if n == "":
        return ""

    if n.startswith("G") or n.startswith("RF"):
        return "POLIS"

    if n.startswith("T"):
        return "ASKAR"

    return "PENGUNDI AWAL"


def is_polis(value):

    n = clean_service_no(value)

    return (
        n.startswith("G")
        or n.startswith("RF")
    )


def is_askar(value):

    n = clean_service_no(value)

    return n.startswith("T")


# ============================================================
# AGE GROUP
# ============================================================

def parse_age_group_ranges(age_groups):
    """
    Convert user-entered age labels into numeric ranges.

    Supported examples:

    18-21
    22-30
    31-40
    61+

    The labels themselves are preserved exactly in the output.
    """

    parsed = []

    for label in age_groups:

        label = str(label).strip()

        # 61+
        match_plus = re.fullmatch(
            r"(\d+)\s*\+",
            label
        )

        if match_plus:

            start = int(match_plus.group(1))

            parsed.append(
                (
                    label,
                    start,
                    None
                )
            )

            continue

        # 18-21
        match_range = re.fullmatch(
            r"(\d+)\s*-\s*(\d+)",
            label
        )

        if match_range:

            start = int(match_range.group(1))
            end = int(match_range.group(2))

            if end < start:
                raise ValueError(
                    f"Invalid age group: {label}"
                )

            parsed.append(
                (
                    label,
                    start,
                    end
                )
            )

            continue

        raise ValueError(
            f"Invalid age group format: '{label}'. "
            f"Use formats such as 18-21 or 61+."
        )

    return parsed


def get_age_group(value, parsed_age_groups):

    if pd.isna(value):
        return ""

    try:
        age = int(float(value))
    except Exception:
        return ""

    for label, start, end in parsed_age_groups:

        if end is None:

            if age >= start:
                return label

        else:

            if start <= age <= end:
                return label

    return ""


# ============================================================
# PERCENTAGE
# ============================================================

def pct(part, total):

    if not total:
        return 0

    return round(
        (part / total) * 100,
        1
    )


# ============================================================
# BUILD DM ROW
# ============================================================

def build_dm_row(
    kod_dm,
    nama_dm,
    grp,
    age_groups
):

    total = len(grp)

    row = {
        "KOD DM": format_kod_dm(kod_dm),
        "NAMA DM": nama_dm,
        "JUMLAH": total
    }

    # --------------------------------------------------------
    # SEX
    # --------------------------------------------------------

    sex_vc = grp["_jantina"].value_counts()

    for key, label in [
        ("L", "LELAKI"),
        ("P", "PEREMPUAN")
    ]:

        count = sex_vc.get(key, 0)

        row[label] = count

        row[f"{label} (%)"] = pct(
            count,
            total
        )

    # --------------------------------------------------------
    # RACE
    # --------------------------------------------------------

    race_vc = grp["_race"].value_counts()

    for race in MAIN_RACES:

        count = race_vc.get(
            race,
            0
        )

        row[race] = count

        row[f"{race} (%)"] = pct(
            count,
            total
        )

    # --------------------------------------------------------
    # AGE
    # --------------------------------------------------------

    age_vc = grp["_age_group"].value_counts()

    for age_group in age_groups:

        count = age_vc.get(
            age_group,
            0
        )

        row[age_group] = count

        row[f"{age_group} (%)"] = pct(
            count,
            total
        )

    # --------------------------------------------------------
    # PARTY
    # --------------------------------------------------------

    party_vc = grp["_party"].value_counts()

    for party in PARTY_COLS:

        count = party_vc.get(
            party,
            0
        )

        row[party] = count

        row[f"{party} (%)"] = pct(
            count,
            total
        )

    # --------------------------------------------------------
    # SIKAP
    # --------------------------------------------------------

    sikap_vc = grp["_sikap"].value_counts()

    for sikap in SIKAP_COLS:

        count = sikap_vc.get(
            sikap,
            0
        )

        row[sikap] = count

        row[f"{sikap} (%)"] = pct(
            count,
            total
        )

    # --------------------------------------------------------
    # PENGUNDI AWAL
    # --------------------------------------------------------

    awal_vc = grp["_awal_type"].value_counts()

    pengundi_awal = (
        grp["_NoPerkhidmatan_clean"]
        .ne("")
        .sum()
    )

    polis = awal_vc.get(
        "POLIS",
        0
    )

    askar = awal_vc.get(
        "ASKAR",
        0
    )

    pasangan_polis = int(
        grp["_Pasangan Polis"].sum()
    )

    pasangan_askar = int(
        grp["_Pasangan Askar"].sum()
    )

    row["PENGUNDI AWAL"] = pengundi_awal

    row["PENGUNDI AWAL (%)"] = pct(
        pengundi_awal,
        total
    )

    row["POLIS"] = polis

    row["POLIS (%)"] = pct(
        polis,
        total
    )

    row["PASANGAN POLIS"] = pasangan_polis

    row["PASANGAN POLIS (%)"] = pct(
        pasangan_polis,
        total
    )

    row["ASKAR"] = askar

    row["ASKAR (%)"] = pct(
        askar,
        total
    )

    row["PASANGAN ASKAR"] = pasangan_askar

    row["PASANGAN ASKAR (%)"] = pct(
        pasangan_askar,
        total
    )

    return row


# ============================================================
# TOTAL ROW
# ============================================================

def add_total_row(df):

    total = (
        pd.to_numeric(
            df["JUMLAH"],
            errors="coerce"
        )
        .fillna(0)
        .sum()
    )

    total_row = {
        "KOD DM": "",
        "NAMA DM": "",
        "JUMLAH": total
    }

    for header in HEADERS:

        if header in {
            "KOD DM",
            "NAMA DM",
            "JUMLAH"
        }:
            continue

        if header.endswith("(%)"):

            base = header.replace(
                " (%)",
                ""
            )

            base_total = (
                pd.to_numeric(
                    df[base],
                    errors="coerce"
                )
                .fillna(0)
                .sum()
            )

            total_row[header] = pct(
                base_total,
                total
            )

        else:

            total_row[header] = (
                pd.to_numeric(
                    df[header],
                    errors="coerce"
                )
                .fillna(0)
                .sum()
            )

    return pd.concat(
        [
            df,
            pd.DataFrame([total_row])
        ],
        ignore_index=True
    )


# ============================================================
# BUILD RUMUSAN
# ============================================================

def build_rumusan_df(
    dun_df,
    age_groups
):

    rows = []

    grouped = dun_df.groupby(
        [
            "_KOD DM",
            "_NAMA DM"
        ],
        dropna=False
    )

    for (
        kod_dm,
        nama_dm
    ), grp in grouped:

        rows.append(
            build_dm_row(
                kod_dm,
                nama_dm,
                grp,
                age_groups
            )
        )

    rumusan_df = pd.DataFrame(
        rows
    )

    # Make sure every required column exists.
    for header in HEADERS:

        if header not in rumusan_df.columns:

            rumusan_df[header] = 0

    rumusan_df = rumusan_df[
        HEADERS
    ]

    rumusan_df = rumusan_df.sort_values(
        by="KOD DM",
        kind="stable"
    )

    rumusan_df = add_total_row(
        rumusan_df
    )

    return rumusan_df


# ============================================================
# WRITE WORKSHEET
# ============================================================

def write_dun_sheet(
    ws,
    rumusan_df,
    age_groups
):

    # --------------------------------------------------------
    # Colours
    # --------------------------------------------------------

    BLUE = "9DC3E6"
    GREEN = "A9D18E"
    ORANGE = "F4B183"
    YELLOW = "FFD966"
    PURPLE = "B4A7D6"
    WHITE = "D9D9D9"

    thin = Side(
        style="thin",
        color="000000"
    )

    medium = Side(
        style="medium",
        color="000000"
    )

    # --------------------------------------------------------
    # Column groups
    # --------------------------------------------------------

    group_fill = {}

    # LELAKI / PEREMPUAN
    for c in range(4, 8):
        group_fill[c] = BLUE

    # RACE
    for c in range(8, 16):
        group_fill[c] = GREEN

    # AGE
    for c in range(16, 28):
        group_fill[c] = ORANGE

    # PARTY
    for c in range(28, 36):
        group_fill[c] = YELLOW

    # SIKAP
    for c in range(36, 42):
        group_fill[c] = WHITE

    # AWAL / POLIS / ASKAR
    for c in range(42, 52):
        group_fill[c] = PURPLE

    group_left_edges = {
        1,
        4,
        8,
        16,
        28,
        36,
        42,
        44,
        46,
        48,
        50
    }

    group_right_edges = {
        2,
        7,
        15,
        27,
        35,
        41,
        43,
        45,
        47,
        49,
        51
    }

    thin_right_edges = {
        5,
        9,
        11,
        13,
        17,
        19,
        21,
        23,
        25,
        29,
        31,
        33,
        37,
        39
    }

    thin_left_edges = {
        10,
        12,
        14,
        18,
        20,
        22,
        24,
        26,
        30,
        32,
        34,
        38,
        40
    }

    # --------------------------------------------------------
    # Header
    # --------------------------------------------------------

    for col_idx, header in enumerate(
        HEADERS,
        start=1
    ):

        cell = ws.cell(
            row=1,
            column=col_idx,
            value=header
        )

        cell.font = Font(
            name="Calibri",
            size=11,
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        if col_idx in group_fill:

            cell.fill = PatternFill(
                "solid",
                fgColor=group_fill[col_idx]
            )

        left = (
            medium
            if col_idx in group_left_edges
            else thin
        )

        right = (
            medium
            if col_idx in group_right_edges
            else thin
        )

        cell.border = Border(
            left=left,
            right=right,
            top=medium,
            bottom=medium
        )

    # --------------------------------------------------------
    # Data rows
    # --------------------------------------------------------

    total_row_number = len(rumusan_df) + 1

    for r_idx, row in enumerate(
        rumusan_df.itertuples(index=False),
        start=2
    ):

        is_total = (
            r_idx == total_row_number
        )

        for c_idx, value in enumerate(
            row,
            start=1
        ):

            cell = ws.cell(
                row=r_idx,
                column=c_idx,
                value=value
            )

            # NAMA DM left-aligned.
            if c_idx == 2:

                cell.alignment = Alignment(
                    horizontal="left",
                    vertical="center"
                )

            else:

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            cell.font = Font(
                name="Calibri",
                size=11,
                bold=is_total
            )

            if (
                c_idx in group_fill
                and is_total
            ):

                cell.fill = PatternFill(
                    "solid",
                    fgColor=group_fill[c_idx]
                )

            left = (
                medium
                if c_idx in group_left_edges
                else thin
            )

            right = (
                medium
                if c_idx in group_right_edges
                else thin
            )

            cell.border = Border(
                left=left,
                right=right,
                top=medium if is_total else thin,
                bottom=medium if is_total else thin
            )

            col_name = HEADERS[
                c_idx - 1
            ]

            if isinstance(
                value,
                (int, float)
            ):

                if "(%)" in col_name:

                    cell.number_format = "0.0"

                else:

                    cell.number_format = "#,##0"

    # --------------------------------------------------------
    # Column widths
    # --------------------------------------------------------

    widths = {

        "A": 13,
        "B": 25,
        "C": 13,

        "D": 11.3,
        "E": 14.7,
        "F": 17,
        "G": 20.6,

        "H": 13.1,
        "I": 16.6,
        "J": 10,
        "K": 13.4,
        "L": 10.7,
        "M": 14.1,
        "N": 14.6,
        "O": 18.1,

        "P": 10.3,
        "Q": 13.7,
        "R": 10.3,
        "S": 13.7,
        "T": 10.3,
        "U": 13.7,
        "V": 10.3,
        "W": 13.7,
        "X": 10.3,
        "Y": 13.7,
        "Z": 8.6,
        "AA": 12,

        "AB": 9,
        "AC": 12.4,
        "AD": 9,
        "AE": 12.4,
        "AF": 10.9,
        "AG": 14.3,
        "AH": 11.7,
        "AI": 15.1,

        "AJ": 11,
        "AK": 14.4,
        "AL": 12.4,
        "AM": 15.9,
        "AN": 11.6,
        "AO": 15,

        "AP": 21.3,
        "AQ": 24.9,
        "AR": 10.6,
        "AS": 14,
        "AT": 21.4,
        "AU": 25,
        "AV": 11.4,
        "AW": 14.9,
        "AX": 22.4,
        "AY": 26
    }

    for col, width in widths.items():

        ws.column_dimensions[
            col
        ].width = width

    # --------------------------------------------------------
    # Auto-adjust NAMA DM
    # --------------------------------------------------------

    max_len = 0

    for row in range(
        1,
        ws.max_row + 1
    ):

        value = ws.cell(
            row=row,
            column=2
        ).value

        max_len = max(
            max_len,
            len(str(value or ""))
        )

    ws.column_dimensions["B"].width = min(
        max_len + 4,
        60
    )

    for row in range(
        2,
        ws.max_row + 1
    ):

        ws.cell(
            row=row,
            column=2
        ).alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

    # --------------------------------------------------------
    # Worksheet settings
    # --------------------------------------------------------

    ws.row_dimensions[1].height = 15.75

    ws.freeze_panes = "A2"

    # --------------------------------------------------------
    # Excel table
    # --------------------------------------------------------

    end_row = len(rumusan_df) + 1

    table_ref = (
        f"A1:AY{end_row}"
    )

    safe_title = re.sub(
        r"[^A-Za-z0-9_]",
        "_",
        ws.title
    )

    table_name = (
        f"Table_{safe_title[:20]}"
        f"_{ws.parent.index(ws)}"
    )

    tab = Table(
        displayName=table_name,
        ref=table_ref
    )

    style = TableStyleInfo(
        name="TableStyleLight1",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=False,
        showColumnStripes=False
    )

    tab.tableStyleInfo = style

    ws.add_table(tab)


# ============================================================
# MAIN GENERATOR
# ============================================================

def generate_demografik(
    uploaded_files,
    age_groups=None
):

    # --------------------------------------------------------
    # Age groups
    # --------------------------------------------------------

    if age_groups is None:

        age_groups = DEFAULT_AGE_GROUPS.copy()

    age_groups = [
        str(x).strip()
        for x in age_groups
    ]

    if len(age_groups) != 6:

        raise ValueError(
            "Exactly 6 age groups are required."
        )

    if any(
        not age
        for age in age_groups
    ):

        raise ValueError(
            "All 6 age groups must be filled in."
        )

    if len(set(age_groups)) != 6:

        raise ValueError(
            "Age groups must be unique."
        )

    parsed_age_groups = parse_age_group_ranges(
        age_groups
    )

    # --------------------------------------------------------
    # Dynamic output headers
    #
    # Replace default age labels with user labels.
    # --------------------------------------------------------

    output_headers = []

    for header in HEADERS:

        replaced = header

        for default_age in DEFAULT_AGE_GROUPS:

            if header == default_age:

                replaced = age_groups[
                    DEFAULT_AGE_GROUPS.index(
                        default_age
                    )
                ]

            elif header == f"{default_age} (%)":

                replaced = (
                    f"{age_groups[DEFAULT_AGE_GROUPS.index(default_age)]}"
                    " (%)"
                )

        output_headers.append(
            replaced
        )

    # --------------------------------------------------------
    # Load files
    # --------------------------------------------------------

    all_data = []
    logs = []

    for uploaded_file in uploaded_files:

        fname = uploaded_file.name

        try:

            # Reset uploaded file pointer.
            uploaded_file.seek(0)

            df = pd.read_excel(
                uploaded_file,
                dtype=str
            )

            df.columns = [
                str(c).strip()
                for c in df.columns
            ]

            # ------------------------------------------------
            # Locate columns
            # ------------------------------------------------

            col_dm = get_col(
                df,
                [
                    "KOD DM",
                    "kod_dm"
                ]
            )

            col_nama_dm = get_col(
                df,
                [
                    "NamaDM",
                    "nama_dm",
                    "NAMA DM"
                ]
            )

            col_kod_dun = get_col(
                df,
                [
                    "kod_dun",
                    "KOD DUN",
                    "KODDUN"
                ]
            )

            col_nama_dun = get_col(
                df,
                [
                    "nama_dun",
                    "DUN",
                    "NAMA DUN"
                ]
            )

            col_jantina = get_col(
                df,
                [
                    "JANTINA",
                    "jantina"
                ]
            )

            col_bangsa = get_col(
                df,
                [
                    "kaum",
                    "BANGSA",
                    "kategori_kaum"
                ]
            )

            col_umur = get_col(
                df,
                [
                    "UMUR",
                    "umur"
                ]
            )

            col_party = get_col(
                df,
                [
                    "party",
                    "PARTY"
                ]
            )

            col_sikap = get_col(
                df,
                [
                    "CATATAN",
                    "sikap"
                ]
            )

            col_no = get_col(
                df,
                [
                    "NoPerkhidmatan",
                    "noperkhidmatan"
                ]
            )

            col_pasangan = get_col(
                df,
                [
                    "NoKPPasangan",
                    "NoPerkhidmatanPasangan",
                    "noperkhidmatanpasangan"
                ]
            )

            # ------------------------------------------------
            # Required columns
            # ------------------------------------------------

            required = {

                "KOD DM": col_dm,

                "NamaDM": col_nama_dm,

                "kod_dun": col_kod_dun,

                "nama_dun / DUN": col_nama_dun,

                "JANTINA": col_jantina,

                "BANGSA": col_bangsa,

                "UMUR": col_umur,

                "NoPerkhidmatan": col_no,

                "NoKPPasangan": col_pasangan
            }

            missing = [
                key
                for key, value in required.items()
                if value is None
            ]

            if missing:

                logs.append(
                    f"Skipped {fname} — "
                    f"missing columns: {missing}"
                )

                continue

            # ------------------------------------------------
            # Internal columns
            # ------------------------------------------------

            df["_KOD_DUN"] = (
                df[col_kod_dun]
                .fillna("")
                .astype(str)
                .str.strip()
            )

            df["_NAMA_DUN"] = (
                df[col_nama_dun]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.upper()
            )

            df["_KOD DM"] = (
                df[col_dm]
                .fillna("")
                .astype(str)
                .str.strip()
            )

            df["_NAMA DM"] = (
                df[col_nama_dm]
                .fillna("")
                .astype(str)
                .str.strip()
            )

            df["_jantina"] = (
                df[col_jantina]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.upper()
            )

            df["_race"] = (
                df[col_bangsa]
                .apply(normalise_race)
            )

            df["_age_group"] = (
                df[col_umur]
                .apply(
                    lambda value:
                    get_age_group(
                        value,
                        parsed_age_groups
                    )
                )
            )

            if col_party:

                df["_party"] = (
                    df[col_party]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

            else:

                df["_party"] = ""

            if col_sikap:

                df["_sikap"] = (
                    df[col_sikap]
                    .apply(normalise_sikap)
                )

            else:

                df["_sikap"] = ""

            # ------------------------------------------------
            # Service numbers
            # ------------------------------------------------

            df["_NoPerkhidmatan_clean"] = (
                df[col_no]
                .apply(clean_service_no)
            )

            df["_NoKPPasangan_clean"] = (
                df[col_pasangan]
                .apply(clean_service_no)
            )

            # ------------------------------------------------
            # Classification
            # ------------------------------------------------

            df["_awal_type"] = (
                df["_NoPerkhidmatan_clean"]
                .apply(classify_awal)
            )

            df["_Pasangan Polis"] = (
                df["_NoKPPasangan_clean"]
                .apply(
                    lambda x:
                    1 if is_polis(x)
                    else 0
                )
            )

            df["_Pasangan Askar"] = (
                df["_NoKPPasangan_clean"]
                .apply(
                    lambda x:
                    1 if is_askar(x)
                    else 0
                )
            )

            all_data.append(df)

            logs.append(
                f"Loaded {fname}: "
                f"{len(df):,} rows"
            )

        except Exception as e:

            logs.append(
                f"Error reading {fname}: {e}"
            )

    # --------------------------------------------------------
    # Validate input
    # --------------------------------------------------------

    if not all_data:

        raise ValueError(
            "No valid data loaded.\n"
            + "\n".join(logs)
        )

    # --------------------------------------------------------
    # Combine all uploaded files
    # --------------------------------------------------------

    final_df = pd.concat(
        all_data,
        ignore_index=True
    )

    # --------------------------------------------------------
    # Find DUNs
    # --------------------------------------------------------

    dun_combos = (
        final_df[
            [
                "_KOD_DUN",
                "_NAMA_DUN"
            ]
        ]
        .drop_duplicates()
    )

    dun_combos = dun_combos[
        (dun_combos["_KOD_DUN"] != "")
        &
        (dun_combos["_NAMA_DUN"] != "")
    ]

    dun_combos = sorted(
        dun_combos.itertuples(
            index=False,
            name=None
        ),
        key=lambda pair:
        kod_dun_sort_key(pair[0])
    )

    if not dun_combos:

        raise ValueError(
            "No DUN name/kod_dun found "
            "in the uploaded data."
        )

    # --------------------------------------------------------
    # Create workbook
    # --------------------------------------------------------

    wb = Workbook()

    wb.remove(
        wb.active
    )

    existing_sheet_names = set()

    # --------------------------------------------------------
    # Build one worksheet per DUN
    # --------------------------------------------------------

    for kod_dun, nama_dun in dun_combos:

        dun_df = final_df[
            (final_df["_KOD_DUN"] == kod_dun)
            &
            (final_df["_NAMA_DUN"] == nama_dun)
        ]

        rumusan_df = build_rumusan_df(
            dun_df,
            age_groups
        )

        # Rename age columns in the final output
        # if custom age labels were supplied.
        rename_map = {}

        for default_age, custom_age in zip(
            DEFAULT_AGE_GROUPS,
            age_groups
        ):

            rename_map[
                default_age
            ] = custom_age

            rename_map[
                f"{default_age} (%)"
            ] = f"{custom_age} (%)"

        rumusan_df = rumusan_df.rename(
            columns=rename_map
        )

        # Make sure output order follows dynamic headers.
        rumusan_df = rumusan_df[
            output_headers
        ]

        sheet_label = format_sheet_label(
            kod_dun,
            nama_dun
        )

        sheet_name = clean_sheet_name(
            sheet_label,
            existing_sheet_names
        )

        ws = wb.create_sheet(
            title=sheet_name
        )

        # Temporarily use dynamic headers inside
        # the worksheet writer.
        original_headers = HEADERS[:]

        try:

            HEADERS[:] = output_headers

            write_dun_sheet(
                ws,
                rumusan_df,
                age_groups
            )

        finally:

            HEADERS[:] = original_headers

        logs.append(
            f"Built sheet '{sheet_name}': "
            f"{len(dun_df):,} rows, "
            f"{len(rumusan_df) - 1} DM(s)"
        )

    # --------------------------------------------------------
    # Save workbook to memory
    # --------------------------------------------------------

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    # --------------------------------------------------------
    # Output filename
    # --------------------------------------------------------

    if len(dun_combos) == 1:

        kod_dun, nama_dun = dun_combos[0]

        out_name = (
            "DEMOGRAFIK "
            + clean_filename(
                format_sheet_label(
                    kod_dun,
                    nama_dun
                )
            )
            + ".xlsx"
        )

    else:

        out_name = (
            f"DEMOGRAFIK "
            f"({len(dun_combos)} DUN).xlsx"
        )

    return (
        output.getvalue(),
        out_name,
        logs
    )
