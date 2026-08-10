import io
import re

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    Border,
    Side,
    PatternFill,
    Alignment
)


# ============================================================
# CONSTANTS
# ============================================================

MAIN_RACES = [
    "MELAYU",
    "CINA",
    "INDIA",
    "LAIN-LAIN"
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
# COLUMN HELPERS
# ============================================================

def get_col(df, possible_names):

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
# CLEANING
# ============================================================

def clean_service_no(value):

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

    name = str(value).strip().upper()

    name = re.sub(
        r'[\\/:*?"<>|]',
        " ",
        name
    )

    name = " ".join(name.split())

    return name if name else "OUTPUT"


# ============================================================
# KOD HELPERS
# ============================================================

def digits_only(value):

    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value.lower() in {
        "",
        "nan",
        "none",
        "null"
    }:
        return ""

    return re.sub(
        r"\D",
        "",
        value.split(".")[0]
    )


def numeric_sort_key(value):

    digits = digits_only(value)

    if not digits:
        return -1

    try:
        return int(digits)
    except Exception:
        return -1


def format_kod_dm(value):

    kod = digits_only(value)

    if not kod:
        return ""

    kod = kod.zfill(7)

    return (
        f"{kod[:3]}/"
        f"{kod[3:5]}/"
        f"{kod[5:]}"
    )


def format_kod_dun(value):

    kod = digits_only(value)

    if not kod:
        return ""

    kod = kod.zfill(5)

    return (
        f"{kod[:3]}/"
        f"{kod[3:5]}"
    )


def format_kod_parlimen(value):

    kod = digits_only(value)

    if not kod:
        return ""

    return kod.zfill(3)


# ============================================================
# NORMALISATION
# ============================================================

def normalise_race(value):

    if pd.isna(value):
        return "LAIN-LAIN"

    race = str(value).strip().upper()

    if race in {
        "MELAYU",
        "CINA",
        "INDIA"
    }:
        return race

    return "LAIN-LAIN"


def normalise_sikap(value):

    if pd.isna(value):
        return ""

    sikap = str(value).strip().upper()

    if sikap in {
        "KELABU-LAMA",
        "KELABU-BARU"
    }:
        return "KELABU"

    if sikap in {
        "PUTIH",
        "KELABU",
        "HITAM"
    }:
        return sikap

    return ""


# ============================================================
# SERVICE CLASSIFICATION
# ============================================================

def classify_awal(value):

    n = clean_service_no(value)

    if not n:
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
# AGE
# ============================================================

def parse_age(value):

    try:

        if pd.isna(value):
            return None

        age = int(float(value))

        return age

    except Exception:

        return None


def get_age_group(value, age_groups):

    age = parse_age(value)

    if age is None:
        return ""

    for label in age_groups:

        label_clean = str(label).strip()

        # ----------------------------------------------------
        # Example: 61+
        # ----------------------------------------------------

        plus_match = re.fullmatch(
            r"(\d+)\s*\+",
            label_clean
        )

        if plus_match:

            minimum = int(
                plus_match.group(1)
            )

            if age >= minimum:
                return label_clean

            continue

        # ----------------------------------------------------
        # Example: 18-21
        # ----------------------------------------------------

        range_match = re.fullmatch(
            r"(\d+)\s*-\s*(\d+)",
            label_clean
        )

        if range_match:

            minimum = int(
                range_match.group(1)
            )

            maximum = int(
                range_match.group(2)
            )

            if minimum <= age <= maximum:
                return label_clean

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
# HEADERS
# ============================================================

def build_headers(age_groups):

    headers = [
        "KOD",
        "NAMA",
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
        "LAIN-LAIN (%)"
    ]

    for age in age_groups:

        headers.append(age)
        headers.append(
            f"{age} (%)"
        )

    headers.extend([
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
    ])

    return headers


# ============================================================
# BUILD ONE SUMMARY ROW
# ============================================================

def build_summary_row(
    kod,
    nama,
    grp,
    age_groups
):

    total = len(grp)

    row = {
        "KOD": kod,
        "NAMA": nama,
        "JUMLAH": total
    }

    # ========================================================
    # SEX
    # ========================================================

    sex_vc = grp["_jantina"].value_counts()

    for key, label in [
        ("L", "LELAKI"),
        ("P", "PEREMPUAN")
    ]:

        count = sex_vc.get(
            key,
            0
        )

        row[label] = count

        row[f"{label} (%)"] = pct(
            count,
            total
        )

    # ========================================================
    # RACE
    # ========================================================

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

    # ========================================================
    # AGE
    # ========================================================

    age_vc = grp["_age_group"].value_counts()

    for age in age_groups:

        count = age_vc.get(
            age,
            0
        )

        row[age] = count

        row[f"{age} (%)"] = pct(
            count,
            total
        )

    # ========================================================
    # PARTY
    # ========================================================

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

    # ========================================================
    # SIKAP
    # ========================================================

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

    # ========================================================
    # PENGUNDI AWAL / POLIS / ASKAR
    # ========================================================

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
# BUILD SUMMARY TABLE
# ============================================================

def build_summary_df(
    df,
    group_columns,
    age_groups,
    headers,
    kod_formatter=None
):

    rows = []

    grouped = df.groupby(
        group_columns,
        dropna=False,
        sort=False
    )

    for group_values, grp in grouped:

        if not isinstance(
            group_values,
            tuple
        ):

            group_values = (
                group_values,
            )

        kod = group_values[0]
        nama = group_values[1]

        if pd.isna(kod):
            kod = ""

        if pd.isna(nama):
            nama = ""

        kod = str(kod).strip()
        nama = str(nama).strip()

        if kod_formatter:
            kod = kod_formatter(kod)

        rows.append(
            build_summary_row(
                kod,
                nama,
                grp,
                age_groups
            )
        )

    summary_df = pd.DataFrame(
        rows
    )

    # --------------------------------------------------------
    # Ensure all headers exist
    # --------------------------------------------------------

    for header in headers:

        if header not in summary_df.columns:
            summary_df[header] = 0

    summary_df = summary_df[headers]

    # --------------------------------------------------------
    # Sort by KOD
    # --------------------------------------------------------

    if not summary_df.empty:

        summary_df["_sort"] = (
            summary_df["KOD"]
            .apply(numeric_sort_key)
        )

        summary_df = (
            summary_df
            .sort_values(
                "_sort",
                kind="stable"
            )
            .drop(
                columns="_sort"
            )
            .reset_index(
                drop=True
            )
        )

    # --------------------------------------------------------
    # Total row
    # --------------------------------------------------------

    summary_df = add_total_row(
        summary_df,
        headers
    )

    return summary_df


# ============================================================
# DUN + DM COMBINED SUMMARY
# ============================================================

def build_dm_with_dun_totals(
    df,
    age_groups,
    headers
):
    """
    Creates ONE worksheet containing:

        DM rows
        DUN subtotal
        blank row
        next DUN DM rows
        next DUN subtotal
        ...

    DUN subtotal rows are calculated directly from the
    underlying records, NOT by averaging DM percentages.
    """

    output_rows = []

    # ========================================================
    # GROUP BY DUN
    # ========================================================

    dun_groups = df.groupby(
        [
            "_KOD DUN",
            "_NAMA DUN"
        ],
        dropna=False,
        sort=False
    )

    dun_blocks = []

    for (
        dun_values,
        dun_grp
    ) in dun_groups:

        if not isinstance(
            dun_values,
            tuple
        ):
            dun_values = (
                dun_values,
            )

        dun_kod = dun_values[0]
        dun_name = dun_values[1]

        if pd.isna(dun_kod):
            dun_kod = ""

        if pd.isna(dun_name):
            dun_name = ""

        dun_kod = str(
            dun_kod
        ).strip()

        dun_name = str(
            dun_name
        ).strip()

        # ----------------------------------------------------
        # DM rows inside this DUN
        # ----------------------------------------------------

        dm_rows = []

        dm_groups = dun_grp.groupby(
            [
                "_KOD DM",
                "_NAMA DM"
            ],
            dropna=False,
            sort=False
        )

        for (
            dm_values,
            dm_grp
        ) in dm_groups:

            if not isinstance(
                dm_values,
                tuple
            ):
                dm_values = (
                    dm_values,
                )

            dm_kod = dm_values[0]
            dm_name = dm_values[1]

            if pd.isna(dm_kod):
                dm_kod = ""

            if pd.isna(dm_name):
                dm_name = ""

            dm_kod = str(
                dm_kod
            ).strip()

            dm_name = str(
                dm_name
            ).strip()

            dm_row = build_summary_row(
                format_kod_dm(dm_kod),
                dm_name,
                dm_grp,
                age_groups
            )

            dm_rows.append(
                dm_row
            )

        # ----------------------------------------------------
        # Sort DM numerically
        # ----------------------------------------------------

        dm_rows.sort(
            key=lambda row:
            numeric_sort_key(
                row["KOD"]
            )
        )

        # ----------------------------------------------------
        # DUN subtotal
        #
        # IMPORTANT:
        # Calculated from ALL underlying records in dun_grp.
        # Percentages therefore use the DUN total as denominator.
        # ----------------------------------------------------

        dun_total_row = build_summary_row(
            format_kod_dun(dun_kod),
            dun_name,
            dun_grp,
            age_groups
        )

        dun_blocks.append(
            (
                numeric_sort_key(
                    format_kod_dun(
                        dun_kod
                    )
                ),
                dm_rows,
                dun_total_row
            )
        )

    # ========================================================
    # SORT DUNs
    # ========================================================

    dun_blocks.sort(
        key=lambda item: item[0]
    )

    # ========================================================
    # CREATE FINAL ROW LIST
    # ========================================================

    for block_index, (
        _sort_key,
        dm_rows,
        dun_total_row
    ) in enumerate(dun_blocks):

        # ----------------------------------------------------
        # DM rows
        # ----------------------------------------------------

        for row in dm_rows:

            output_rows.append(
                row
            )

        # ----------------------------------------------------
        # DUN subtotal
        # ----------------------------------------------------

        output_rows.append(
            dun_total_row
        )

        # ----------------------------------------------------
        # Blank row between DUNs
        # ----------------------------------------------------

        if block_index < len(dun_blocks) - 1:

            output_rows.append(
                {
                    header: ""
                    for header in headers
                }
            )

    # ========================================================
    # DATAFRAME
    # ========================================================

    if output_rows:

        result = pd.DataFrame(
            output_rows
        )

    else:

        result = pd.DataFrame(
            columns=headers
        )

    # --------------------------------------------------------
    # Ensure all headers exist
    # --------------------------------------------------------

    for header in headers:

        if header not in result.columns:
            result[header] = 0

    result = result[headers]

    return result


# ============================================================
# TOTAL ROW
# ============================================================

def add_total_row(
    df,
    headers
):

    if df.empty:

        total = 0

    else:

        total = pd.to_numeric(
            df["JUMLAH"],
            errors="coerce"
        ).fillna(0).sum()

    total_row = {
        "KOD": "",
        "NAMA": "",
        "JUMLAH": total
    }

    for header in headers:

        if header in {
            "KOD",
            "NAMA",
            "JUMLAH"
        }:
            continue

        if header.endswith("(%)"):

            base = header.replace(
                " (%)",
                ""
            )

            numerator = total_row.get(
                base,
                0
            )

            total_row[header] = pct(
                numerator,
                total
            )

        else:

            if header in df.columns:

                total_row[header] = (
                    pd.to_numeric(
                        df[header],
                        errors="coerce"
                    )
                    .fillna(0)
                    .sum()
                )

            else:

                total_row[header] = 0

    return pd.concat(
        [
            df,
            pd.DataFrame(
                [total_row]
            )
        ],
        ignore_index=True
    )


# ============================================================
# SHEET FORMATTING
# ============================================================

def write_summary_sheet(
    ws,
    summary_df,
    headers
):

    # ========================================================
    # COLOURS
    # ========================================================

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

    # ========================================================
    # COLUMN GROUP COLOURS
    # ========================================================

    group_fill = {}

    # --------------------------------------------------------
    # Sex
    # --------------------------------------------------------

    for c in range(4, 8):
        group_fill[c] = BLUE

    # --------------------------------------------------------
    # Race
    # --------------------------------------------------------

    for c in range(8, 16):
        group_fill[c] = GREEN

    # --------------------------------------------------------
    # Age
    #
    # Dynamic based on configured age_groups.
    # Age starts immediately after column O.
    # --------------------------------------------------------

    age_start = 16

    age_count = 0

    for header in headers:

        if header.endswith(" (%)"):

            base = header.replace(
                " (%)",
                ""
            )

            if base not in {
                "LELAKI",
                "PEREMPUAN",
                "MELAYU",
                "CINA",
                "INDIA",
                "LAIN-LAIN",
                "UMNO",
                "PKR",
                "PAS",
                "PPBM",
                "PUTIH",
                "KELABU",
                "HITAM",
                "PENGUNDI AWAL",
                "POLIS",
                "PASANGAN POLIS",
                "ASKAR",
                "PASANGAN ASKAR"
            }:
                age_count += 1

    age_end = (
        age_start
        + (age_count * 2)
        - 1
    )

    for c in range(
        age_start,
        age_end + 1
    ):
        group_fill[c] = ORANGE

    # ========================================================
    # DYNAMIC GROUP LOCATIONS
    # ========================================================

    def find_column(name):

        try:
            return (
                headers.index(name)
                + 1
            )

        except ValueError:

            return None

    party_start = find_column(
        "UMNO"
    )

    if party_start:

        for c in range(
            party_start,
            party_start + 8
        ):
            group_fill[c] = YELLOW

    sikap_start = find_column(
        "PUTIH"
    )

    if sikap_start:

        for c in range(
            sikap_start,
            sikap_start + 6
        ):
            group_fill[c] = WHITE

    awal_start = find_column(
        "PENGUNDI AWAL"
    )

    if awal_start:

        for c in range(
            awal_start,
            len(headers) + 1
        ):
            group_fill[c] = PURPLE

    # ========================================================
    # GROUP BORDERS
    # ========================================================

    group_left_edges = {
        1,
        4,
        8,
        16
    }

    group_right_edges = {
        2,
        7,
        15
    }

    if party_start:

        group_left_edges.add(
            party_start
        )

        group_right_edges.add(
            party_start + 7
        )

    if sikap_start:

        group_left_edges.add(
            sikap_start
        )

        group_right_edges.add(
            sikap_start + 5
        )

    if awal_start:

        group_left_edges.add(
            awal_start
        )

        group_right_edges.add(
            len(headers)
        )

    # ========================================================
    # HEADER
    # ========================================================

    for col_idx, header in enumerate(
        headers,
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

    # ========================================================
    # DATA
    # ========================================================

    for row_idx, row in enumerate(
        summary_df.itertuples(
            index=False
        ),
        start=2
    ):

        values = list(row)

        # ----------------------------------------------------
        # Detect row type
        # ----------------------------------------------------

        kod_value = values[0]
        nama_value = values[1]

        is_blank = (
            str(kod_value or "") == ""
            and
            str(nama_value or "") == ""
            and
            all(
                str(v or "") == ""
                for v in values[2:]
            )
        )

        is_dun_total = False

        if not is_blank:

            kod_digits = digits_only(
                kod_value
            )

            # DUN code has 5 digits:
            # 082/12
            #
            # DM code has 7 digits:
            # 082/12/01
            if (
                kod_digits
                and len(kod_digits) <= 5
            ):
                is_dun_total = True

        # ----------------------------------------------------
        # Write cells
        # ----------------------------------------------------

        for col_idx, value in enumerate(
            values,
            start=1
        ):

            cell = ws.cell(
                row=row_idx,
                column=col_idx,
                value=value
            )

            # ------------------------------------------------
            # Blank separator row
            # ------------------------------------------------

            if is_blank:

                cell.border = Border()

                continue

            # ------------------------------------------------
            # Alignment
            # ------------------------------------------------

            if col_idx == 2:

                cell.alignment = Alignment(
                    horizontal="left",
                    vertical="center"
                )

            else:

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            # ------------------------------------------------
            # Font
            # ------------------------------------------------

            cell.font = Font(
                name="Calibri",
                size=11,
                bold=is_dun_total
            )

            # ------------------------------------------------
            # DUN subtotal fill
            #
            # Keep the existing group colour scheme but make
            # the subtotal row bold.
            # ------------------------------------------------

            if (
                is_dun_total
                and col_idx in group_fill
            ):

                cell.fill = PatternFill(
                    "solid",
                    fgColor=group_fill[col_idx]
                )

            elif col_idx in group_fill:

                cell.fill = PatternFill(
                    "solid",
                    fgColor=group_fill[col_idx]
                )

            # ------------------------------------------------
            # Borders
            # ------------------------------------------------

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
                top=(
                    medium
                    if is_dun_total
                    else thin
                ),
                bottom=(
                    medium
                    if is_dun_total
                    else thin
                )
            )

            # ------------------------------------------------
            # Number format
            # ------------------------------------------------

            header = headers[
                col_idx - 1
            ]

            if isinstance(
                value,
                (int, float)
            ):

                if header.endswith(
                    "(%)"
                ):

                    cell.number_format = "0.0"

                else:

                    cell.number_format = "#,##0"

    # ========================================================
    # COLUMN WIDTHS
    # ========================================================

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

    for column, width in widths.items():

        ws.column_dimensions[
            column
        ].width = width

    # ========================================================
    # DYNAMIC WIDTH FOR NAMA
    # ========================================================

    max_len = 0

    for row_idx in range(
        1,
        ws.max_row + 1
    ):

        value = ws.cell(
            row=row_idx,
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

    # ========================================================
    # ROW HEIGHT
    # ========================================================

    ws.row_dimensions[1].height = 15.75

    for row_idx in range(
        2,
        ws.max_row + 1
    ):

        # Blank separator rows are slightly smaller
        is_blank = all(
            ws.cell(
                row=row_idx,
                column=col
            ).value in (
                None,
                ""
            )
            for col in range(
                1,
                len(headers) + 1
            )
        )

        if is_blank:

            ws.row_dimensions[
                row_idx
            ].height = 8

        else:

            ws.row_dimensions[
                row_idx
            ].height = 15

    # ========================================================
    # GENERAL WORKSHEET SETTINGS
    # ========================================================

    ws.freeze_panes = "A2"

    # ========================================================
    # PAGE SETUP
    # ========================================================

    ws.page_setup.orientation = "landscape"

    ws.page_setup.paperSize = (
        ws.PAPERSIZE_A4
    )

    ws.page_setup.fitToWidth = 1

    ws.page_setup.fitToHeight = 0

    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2


# ============================================================
# PREPARE INPUT DATA
# ============================================================

def prepare_dataframe(
    df,
    age_groups
):

    df = df.copy()

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    # ========================================================
    # PARLIAMENT
    # ========================================================

    col_kod_parlimen = get_col(
        df,
        [
            "kod_parlimen",
            "KOD PARLIMEN",
            "KODPARLIMEN"
        ]
    )

    col_nama_parlimen = get_col(
        df,
        [
            "nama_parlimen",
            "NamaParlimen",
            "NAMA PARLIMEN"
        ]
    )

    # ========================================================
    # DUN
    # ========================================================

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

    # ========================================================
    # DM
    # ========================================================

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

    # ========================================================
    # OTHER COLUMNS
    # ========================================================

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

    required = {
        "kod_parlimen": col_kod_parlimen,
        "nama_parlimen": col_nama_parlimen,
        "kod_dun": col_kod_dun,
        "nama_dun / DUN": col_nama_dun,
        "KOD DM": col_dm,
        "NamaDM": col_nama_dm,
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

        raise ValueError(
            "Missing required columns: "
            + ", ".join(missing)
        )

    # ========================================================
    # HIERARCHY
    # ========================================================

    df["_KOD PARLIMEN"] = (
        df[col_kod_parlimen]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["_NAMA PARLIMEN"] = (
        df[col_nama_parlimen]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df["_KOD DUN"] = (
        df[col_kod_dun]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["_NAMA DUN"] = (
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

    # ========================================================
    # DEMOGRAPHIC FIELDS
    # ========================================================

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

    # IMPORTANT:
    # Age groups come directly from the Streamlit inputs.
    df["_age_group"] = (
        df[col_umur]
        .apply(
            lambda x:
            get_age_group(
                x,
                age_groups
            )
        )
    )

    # ========================================================
    # PARTY
    # ========================================================

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

    # ========================================================
    # SIKAP
    # ========================================================

    if col_sikap:

        df["_sikap"] = (
            df[col_sikap]
            .apply(normalise_sikap)
        )

    else:

        df["_sikap"] = ""

    # ========================================================
    # SERVICE NUMBERS
    # ========================================================

    df["_NoPerkhidmatan_clean"] = (
        df[col_no]
        .apply(clean_service_no)
    )

    df["_NoKPPasangan_clean"] = (
        df[col_pasangan]
        .apply(clean_service_no)
    )

    # ========================================================
    # CLASSIFICATION
    # ========================================================

    df["_awal_type"] = (
        df["_NoPerkhidmatan_clean"]
        .apply(classify_awal)
    )

    df["_Pasangan Polis"] = (
        df["_NoKPPasangan_clean"]
        .apply(
            lambda x:
            1 if is_polis(x) else 0
        )
    )

    df["_Pasangan Askar"] = (
        df["_NoKPPasangan_clean"]
        .apply(
            lambda x:
            1 if is_askar(x) else 0
        )
    )

    return df


# ============================================================
# GENERATE
# ============================================================

def generate_demografik(
    uploaded_files,
    age_groups
):

    logs = []

    all_data = []

    headers = build_headers(
        age_groups
    )

    # ========================================================
    # READ FILES
    # ========================================================

    for uploaded_file in uploaded_files:

        fname = uploaded_file.name

        try:

            df = pd.read_excel(
                uploaded_file,
                dtype=str
            )

            prepared = prepare_dataframe(
                df,
                age_groups
            )

            all_data.append(
                prepared
            )

            logs.append(
                f"Loaded {fname}: "
                f"{len(prepared):,} rows"
            )

        except Exception as e:

            logs.append(
                f"Error reading {fname}: "
                f"{e}"
            )

    if not all_data:

        raise ValueError(
            "No valid data loaded.\n"
            + "\n".join(logs)
        )

    final_df = pd.concat(
        all_data,
        ignore_index=True
    )

    # ========================================================
    # WORKBOOK
    # ========================================================

    wb = Workbook()

    default_ws = wb.active

    wb.remove(
        default_ws
    )

    # ========================================================
    # 1. PARLIMEN
    # ========================================================

    ws_parlimen = wb.create_sheet(
        title="PARLIMEN"
    )

    parlimen_df = build_summary_df(
        final_df,
        [
            "_KOD PARLIMEN",
            "_NAMA PARLIMEN"
        ],
        age_groups,
        headers,
        kod_formatter=format_kod_parlimen
    )

    write_summary_sheet(
        ws_parlimen,
        parlimen_df,
        headers
    )

    logs.append(
        f"PARLIMEN worksheet: "
        f"{len(parlimen_df) - 1:,} Parliament(s)"
    )

    # ========================================================
    # 2. DM
    #
    # ONE worksheet ONLY.
    #
    # DM rows are grouped by DUN.
    # Each DUN gets a subtotal row.
    # ========================================================

    ws_dm = wb.create_sheet(
        title="DM"
    )

    dm_with_dun_df = (
        build_dm_with_dun_totals(
            final_df,
            age_groups,
            headers
        )
    )

    write_summary_sheet(
        ws_dm,
        dm_with_dun_df,
        headers
    )

    # --------------------------------------------------------
    # Count DUNs and DMs for logging
    # --------------------------------------------------------

    dun_count = (
        final_df[
            [
                "_KOD DUN",
                "_NAMA DUN"
            ]
        ]
        .drop_duplicates()
        .shape[0]
    )

    dm_count = (
        final_df[
            [
                "_KOD DM",
                "_NAMA DM"
            ]
        ]
        .drop_duplicates()
        .shape[0]
    )

    logs.append(
        f"DM worksheet: "
        f"{dm_count:,} DM(s), "
        f"{dun_count:,} DUN subtotal(s)"
    )

    logs.append(
        "DUN subtotals calculated directly "
        "from underlying records; percentages "
        "were recalculated from DUN totals."
    )

    logs.append(
        "Age groups used: "
        + ", ".join(age_groups)
    )

    # ========================================================
    # SAVE
    # ========================================================

    output = io.BytesIO()

    wb.save(
        output
    )

    output.seek(0)

    # ========================================================
    # OUTPUT NAME
    # ========================================================

    if len(uploaded_files) == 1:

        source_name = (
            uploaded_files[0].name
        )

        source_name = re.sub(
            r"\.(xlsx|xls)$",
            "",
            source_name,
            flags=re.IGNORECASE
        )

        out_name = (
            f"DEMOGRAFIK "
            f"{clean_filename(source_name)}.xlsx"
        )

    else:

        out_name = (
            "DEMOGRAFIK "
            f"({len(uploaded_files)} FILES).xlsx"
        )

    return (
        output.getvalue(),
        out_name,
        logs
    )
