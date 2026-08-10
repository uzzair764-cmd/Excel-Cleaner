import io
import re

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Border, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo


# ============================================================
# COLUMN DEFINITIONS
# ============================================================

HEADERS = [
    'KOD',
    'NAMA',
    'JUMLAH',

    'LELAKI',
    'LELAKI (%)',
    'PEREMPUAN',
    'PEREMPUAN (%)',

    'MELAYU',
    'MELAYU (%)',
    'CINA',
    'CINA (%)',
    'INDIA',
    'INDIA (%)',
    'LAIN-LAIN',
    'LAIN-LAIN (%)',

    '18-21',
    '18-21 (%)',
    '22-30',
    '22-30 (%)',
    '31-40',
    '31-40 (%)',
    '41-50',
    '41-50 (%)',
    '51-60',
    '51-60 (%)',
    '61+',
    '61+ (%)',

    'UMNO',
    'UMNO (%)',
    'PKR',
    'PKR (%)',
    'PAS',
    'PAS (%)',
    'PPBM',
    'PPBM (%)',

    'PUTIH',
    'PUTIH (%)',
    'KELABU',
    'KELABU (%)',
    'HITAM',
    'HITAM (%)',

    'PENGUNDI AWAL',
    'PENGUNDI AWAL (%)',

    'POLIS',
    'POLIS (%)',

    'PASANGAN POLIS',
    'PASANGAN POLIS (%)',

    'ASKAR',
    'ASKAR (%)',

    'PASANGAN ASKAR',
    'PASANGAN ASKAR (%)',
]

MAIN_RACES = [
    'MELAYU',
    'CINA',
    'INDIA',
    'LAIN-LAIN'
]

AGE_GROUPS = [
    '18-21',
    '22-30',
    '31-40',
    '41-50',
    '51-60',
    '61+'
]

PARTY_COLS = [
    'UMNO',
    'PKR',
    'PAS',
    'PPBM'
]

SIKAP_COLS = [
    'PUTIH',
    'KELABU',
    'HITAM'
]

AWAL_COLS = [
    'PENGUNDI AWAL',
    'POLIS',
    'PASANGAN POLIS',
    'ASKAR',
    'PASANGAN ASKAR'
]


# ============================================================
# COLUMN HELPERS
# ============================================================

def get_col(df, possible_names):
    """
    Find a dataframe column using case-insensitive matching.
    """

    col_map = {
        str(c).strip().lower(): c
        for c in df.columns
    }

    for name in possible_names:
        key = str(name).strip().lower()

        if key in col_map:
            return col_map[key]

    return None


# ============================================================
# GENERAL CLEANING
# ============================================================

def clean_filename(value):
    name = str(value).strip().upper()

    name = re.sub(
        r'[\\/:*?"<>|]',
        ' ',
        name
    )

    name = ' '.join(name.split())

    return name if name else 'OUTPUT'


def clean_service_no(value):
    if pd.isna(value):
        return ''

    value = str(value).strip().upper()

    if value in {
        '',
        'NAN',
        'NONE',
        'NULL',
        'N/A',
        '#N/A'
    }:
        return ''

    return value


# ============================================================
# KOD FORMATTING
# ============================================================

def format_kod_dm(value):
    """
    Converts:

        0821201
        821201
        082/12/01

    into:

        082/12/01
    """

    if pd.isna(value):
        return ''

    kod = str(value).strip()

    if kod in {
        '',
        'NONE',
        'NAN',
        'NULL',
        'N/A',
        '#N/A'
    }:
        return ''

    # Remove Excel-style decimal suffix.
    if '.' in kod:
        kod = kod.split('.')[0]

    # Already formatted.
    if '/' in kod:
        parts = kod.split('/')

        if len(parts) == 3:
            return (
                parts[0].zfill(3)
                + '/'
                + parts[1].zfill(2)
                + '/'
                + parts[2].zfill(2)
            )

    # Digits only.
    kod = re.sub(r'\D', '', kod)

    if not kod:
        return ''

    kod = kod.zfill(7)

    return (
        kod[:3]
        + '/'
        + kod[3:5]
        + '/'
        + kod[5:7]
    )


def kod_dun_from_dm(kod_dm):
    """
    082/12/01 -> 082/12
    """

    kod = format_kod_dm(kod_dm)

    if not kod:
        return ''

    parts = kod.split('/')

    if len(parts) != 3:
        return ''

    return f'{parts[0]}/{parts[1]}'


def dun_sort_key(value):
    """
    Sort DUN numerically where possible.
    """

    value = str(value).strip()

    digits = re.sub(r'\D', '', value)

    if digits:
        try:
            return int(digits)
        except Exception:
            pass

    return 999999


def dm_sort_key(value):
    """
    Sort DM numerically by the final component.
    """

    kod = format_kod_dm(value)

    if not kod:
        return 999999

    parts = kod.split('/')

    try:
        return (
            int(parts[0]),
            int(parts[1]),
            int(parts[2])
        )
    except Exception:
        return (999999, 999999, 999999)


# ============================================================
# NORMALISATION
# ============================================================

def normalise_race(value):

    if pd.isna(value):
        return 'LAIN-LAIN'

    race = str(value).strip().upper()

    if race in {
        'MELAYU',
        'CINA',
        'INDIA'
    }:
        return race

    return 'LAIN-LAIN'


def normalise_sikap(value):

    if pd.isna(value):
        return None

    sikap = str(value).strip().upper()

    if sikap in {
        'KELABU',
        'KELABU-LAMA',
        'KELABU-BARU'
    }:
        return 'KELABU'

    if sikap == 'PUTIH':
        return 'PUTIH'

    if sikap == 'HITAM':
        return 'HITAM'

    return None


# ============================================================
# PENGUNDI AWAL
# ============================================================

def classify_pengundi_awal(value):

    n = clean_service_no(value)

    if not n:
        return None

    if n.startswith('G') or n.startswith('RF'):
        return 'POLIS'

    if n.startswith('T'):
        return 'ASKAR'

    return 'PENGUNDI AWAL'


def is_polis(value):

    n = clean_service_no(value)

    return (
        n.startswith('G')
        or n.startswith('RF')
    )


def is_askar(value):

    n = clean_service_no(value)

    return n.startswith('T')


# ============================================================
# AGE GROUP
# ============================================================

def get_age_group(value):

    if pd.isna(value):
        return None

    try:

        age = int(float(value))

        if 18 <= age <= 21:
            return '18-21'

        if 22 <= age <= 30:
            return '22-30'

        if 31 <= age <= 40:
            return '31-40'

        if 41 <= age <= 50:
            return '41-50'

        if 51 <= age <= 60:
            return '51-60'

        if age >= 61:
            return '61+'

    except Exception:
        pass

    return None


# ============================================================
# PERCENTAGE
# ============================================================

def pct(part, total):

    if not total:
        return 0.0

    return round(
        float(part) / float(total) * 100,
        1
    )


# ============================================================
# BUILD STATISTICS ROW
# ============================================================

def build_stats_row(
    grp,
    kod='',
    nama='',
):
    """
    Build one statistics row.
    """

    # --------------------------------------------------------
    # TOTAL
    # --------------------------------------------------------

    total = len(grp)

    # --------------------------------------------------------
    # VALUE COUNTS
    # --------------------------------------------------------

    race_vc = (
        grp['race_norm']
        .value_counts()
    )

    age_vc = (
        grp['age_group']
        .value_counts()
    )

    party_vc = (
        grp['party_norm']
        .value_counts()
    )

    sikap_vc = (
        grp['sikap_norm']
        .value_counts()
    )

    awal_vc = (
        grp['awal_type']
        .value_counts()
    )

    sex_vc = (
        grp['jantina_norm']
        .value_counts()
    )

    # --------------------------------------------------------
    # BASE
    # --------------------------------------------------------

    row = {
        'KOD': kod,
        'NAMA': nama,
        'JUMLAH': total
    }

    # --------------------------------------------------------
    # JANTINA
    # --------------------------------------------------------

    lelaki = int(
        sex_vc.get('L', 0)
    )

    perempuan = int(
        sex_vc.get('P', 0)
    )

    row['LELAKI'] = lelaki
    row['LELAKI (%)'] = pct(
        lelaki,
        total
    )

    row['PEREMPUAN'] = perempuan
    row['PEREMPUAN (%)'] = pct(
        perempuan,
        total
    )

    # --------------------------------------------------------
    # KAUM
    # --------------------------------------------------------

    for race in MAIN_RACES:

        count = int(
            race_vc.get(
                race,
                0
            )
        )

        row[race] = count
        row[f'{race} (%)'] = pct(
            count,
            total
        )

    # --------------------------------------------------------
    # AGE
    # --------------------------------------------------------

    for age_group in AGE_GROUPS:

        count = int(
            age_vc.get(
                age_group,
                0
            )
        )

        row[age_group] = count
        row[f'{age_group} (%)'] = pct(
            count,
            total
        )

    # --------------------------------------------------------
    # PARTY
    # --------------------------------------------------------

    for party in PARTY_COLS:

        count = int(
            party_vc.get(
                party,
                0
            )
        )

        row[party] = count
        row[f'{party} (%)'] = pct(
            count,
            total
        )

    # --------------------------------------------------------
    # SIKAP
    # --------------------------------------------------------

    for sikap in SIKAP_COLS:

        count = int(
            sikap_vc.get(
                sikap,
                0
            )
        )

        row[sikap] = count
        row[f'{sikap} (%)'] = pct(
            count,
            total
        )

    # --------------------------------------------------------
    # PENGUNDI AWAL
    # --------------------------------------------------------

    pengundi_awal = int(
        grp['NoPerkhidmatan_clean']
        .ne('')
        .sum()
    )

    row['PENGUNDI AWAL'] = pengundi_awal

    row['PENGUNDI AWAL (%)'] = pct(
        pengundi_awal,
        total
    )

    # --------------------------------------------------------
    # POLIS
    # --------------------------------------------------------

    polis = int(
        awal_vc.get(
            'POLIS',
            0
        )
    )

    row['POLIS'] = polis

    row['POLIS (%)'] = pct(
        polis,
        total
    )

    # --------------------------------------------------------
    # PASANGAN POLIS
    # --------------------------------------------------------

    pasangan_polis = int(
        grp['Pasangan Polis']
        .sum()
    )

    row['PASANGAN POLIS'] = pasangan_polis

    row['PASANGAN POLIS (%)'] = pct(
        pasangan_polis,
        total
    )

    # --------------------------------------------------------
    # ASKAR
    # --------------------------------------------------------

    askar = int(
        awal_vc.get(
            'ASKAR',
            0
        )
    )

    row['ASKAR'] = askar

    row['ASKAR (%)'] = pct(
        askar,
        total
    )

    # --------------------------------------------------------
    # PASANGAN ASKAR
    # --------------------------------------------------------

    pasangan_askar = int(
        grp['Pasangan Askar']
        .sum()
    )

    row['PASANGAN ASKAR'] = pasangan_askar

    row['PASANGAN ASKAR (%)'] = pct(
        pasangan_askar,
        total
    )

    return row


# ============================================================
# TOTAL ROW
# ============================================================

def build_total_row(
    df,
    label=''
):
    """
    Creates a total row from a dataframe containing
    statistics rows.
    """

    if df.empty:

        row = {
            header: ''
            for header in HEADERS
        }

        row['NAMA'] = label

        return row

    total = int(
        pd.to_numeric(
            df['JUMLAH'],
            errors='coerce'
        )
        .fillna(0)
        .sum()
    )

    row = {
        header: ''
        for header in HEADERS
    }

    row['NAMA'] = label
    row['JUMLAH'] = total

    # --------------------------------------------------------
    # COUNT COLUMNS
    # --------------------------------------------------------

    count_columns = (
        [
            'LELAKI',
            'PEREMPUAN'
        ]
        + MAIN_RACES
        + AGE_GROUPS
        + PARTY_COLS
        + SIKAP_COLS
        + AWAL_COLS
    )

    for column in count_columns:

        if column not in df.columns:
            row[column] = 0
            continue

        row[column] = int(
            pd.to_numeric(
                df[column],
                errors='coerce'
            )
            .fillna(0)
            .sum()
        )

    # --------------------------------------------------------
    # PERCENTAGE COLUMNS
    # --------------------------------------------------------

    for header in HEADERS:

        if '(%)' not in header:
            continue

        base = header.replace(
            ' (%)',
            ''
        )

        row[header] = pct(
            row.get(base, 0),
            total
        )

    return row


# ============================================================
# GRAND TOTAL
# ============================================================

def build_grand_total(df):

    if df.empty:
        return {
            header: ''
            for header in HEADERS
        }

    return build_total_row(
        df,
        'GRAND TOTAL'
    )


# ============================================================
# EXCEL FORMATTING
# ============================================================

def reset_cell_style(cell):

    cell.font = Font(
        name='Calibri',
        size=11,
        bold=False
    )

    cell.border = Border()

    cell.fill = PatternFill(
        fill_type=None
    )

    cell.alignment = Alignment(
        vertical='center'
    )


def format_worksheet(ws):

    # --------------------------------------------------------
    # Remove all default styling
    # --------------------------------------------------------

    for row in ws.iter_rows():

        for cell in row:

            reset_cell_style(cell)

    # --------------------------------------------------------
    # Header
    # --------------------------------------------------------

    for cell in ws[1]:

        cell.font = Font(
            name='Calibri',
            size=11,
            bold=False
        )

        cell.alignment = Alignment(
            horizontal='center',
            vertical='center',
            wrap_text=True
        )

    # --------------------------------------------------------
    # Data
    # --------------------------------------------------------

    for row in ws.iter_rows(
        min_row=2,
        max_row=ws.max_row
    ):

        for cell in row:

            header = ws.cell(
                row=1,
                column=cell.column
            ).value

            if header == 'NAMA':

                cell.alignment = Alignment(
                    horizontal='left',
                    vertical='center'
                )

            else:

                cell.alignment = Alignment(
                    horizontal='center',
                    vertical='center'
                )

            if isinstance(
                cell.value,
                (int, float)
            ):

                if (
                    isinstance(header, str)
                    and '(%)' in header
                ):

                    cell.number_format = '0.0'

                else:

                    cell.number_format = '#,##0'

    # --------------------------------------------------------
    # Column widths
    # --------------------------------------------------------

    widths = {
        'A': 14,
        'B': 25,
        'C': 13,

        'D': 11,
        'E': 13,
        'F': 14,
        'G': 16,

        'H': 12,
        'I': 13,
        'J': 10,
        'K': 12,
        'L': 10,
        'M': 12,
        'N': 13,
        'O': 15,

        'P': 10,
        'Q': 13,
        'R': 10,
        'S': 13,
        'T': 10,
        'U': 13,
        'V': 10,
        'W': 13,
        'X': 10,
        'Y': 13,
        'Z': 9,
        'AA': 12,

        'AB': 10,
        'AC': 12,
        'AD': 10,
        'AE': 12,
        'AF': 10,
        'AG': 12,
        'AH': 10,
        'AI': 12,

        'AJ': 10,
        'AK': 13,
        'AL': 10,
        'AM': 13,
        'AN': 10,
        'AO': 13,

        'AP': 20,
        'AQ': 24,
        'AR': 10,
        'AS': 13,
        'AT': 20,
        'AU': 24,
        'AV': 10,
        'AW': 13,
        'AX': 20,
        'AY': 24
    }

    for column, width in widths.items():

        ws.column_dimensions[
            column
        ].width = width

    # --------------------------------------------------------
    # Improve NAMA width
    # --------------------------------------------------------

    max_name_length = 0

    for row in range(
        1,
        ws.max_row + 1
    ):

        value = ws.cell(
            row=row,
            column=2
        ).value

        max_name_length = max(
            max_name_length,
            len(str(value or ''))
        )

    ws.column_dimensions['B'].width = min(
        max_name_length + 4,
        40
    )

    # --------------------------------------------------------
    # Row heights
    # --------------------------------------------------------

    ws.row_dimensions[1].height = 30

    for row in range(
        2,
        ws.max_row + 1
    ):

        # Blank separator rows remain compact.
        if all(
            ws.cell(
                row=row,
                column=column
            ).value in ('', None)
            for column in range(
                1,
                ws.max_column + 1
            )
        ):

            ws.row_dimensions[row].height = 8

        else:

            ws.row_dimensions[row].height = 22

    # --------------------------------------------------------
    # Freeze
    # --------------------------------------------------------

    ws.freeze_panes = 'A2'

    # --------------------------------------------------------
    # Auto filter
    # --------------------------------------------------------

    ws.auto_filter.ref = (
        f'A1:AY{ws.max_row}'
    )


# ============================================================
# WRITE DATAFRAME
# ============================================================

def write_dataframe(
    ws,
    df
):

    # Always use exact HEADERS order.
    df = df.reindex(
        columns=HEADERS,
        fill_value=''
    )

    # Header.
    for column_index, header in enumerate(
        HEADERS,
        start=1
    ):

        ws.cell(
            row=1,
            column=column_index,
            value=header
        )

    # Data.
    for row_index, row in enumerate(
        df.itertuples(
            index=False,
            name=None
        ),
        start=2
    ):

        for column_index, value in enumerate(
            row,
            start=1
        ):

            ws.cell(
                row=row_index,
                column=column_index,
                value=value
            )

    format_worksheet(ws)


# ============================================================
# BUILD PARLIMEN DATA
# ============================================================

def build_parlimen_dataframe(
    final_df
):

    rows = []

    grouped = final_df.groupby(
        [
            '_KOD_PARLIMEN',
            '_NAMA_PARLIMEN'
        ],
        dropna=False,
        sort=False
    )

    for (
        kod_parlimen,
        nama_parlimen
    ), grp in grouped:

        row = build_stats_row(
            grp,
            kod=str(
                kod_parlimen
            ).strip(),
            nama=str(
                nama_parlimen
            ).strip()
        )

        rows.append(row)

    if not rows:
        return pd.DataFrame(
            columns=HEADERS
        )

    df = pd.DataFrame(rows)

    # --------------------------------------------------------
    # Sort Parliament
    # --------------------------------------------------------

    df['_SORT'] = (
        df['KOD']
        .astype(str)
        .str.extract(
            r'(\d+)'
        )[0]
        .fillna('999999')
        .astype(int)
    )

    df = (
        df.sort_values(
            '_SORT'
        )
        .drop(
            columns=['_SORT']
        )
        .reset_index(
            drop=True
        )
    )

    # --------------------------------------------------------
    # Grand total
    # --------------------------------------------------------

    grand = build_grand_total(df)

    df = pd.concat(
        [
            df,
            pd.DataFrame([grand])
        ],
        ignore_index=True
    )

    return df.reindex(
        columns=HEADERS
    )


# ============================================================
# BUILD DM DATA GROUPED BY DUN
# ============================================================

def build_dm_grouped_dataframe(
    final_df
):
    """
    Creates:

        DUN 12 DM 01
        DUN 12 DM 02
        DUN 12 TOTAL
        blank

        DUN 13 DM 01
        DUN 13 DM 02
        DUN 13 TOTAL
        blank

        ...

        GRAND TOTAL
    """

    output_rows = []

    # --------------------------------------------------------
    # Get unique DUNs
    # --------------------------------------------------------

    dun_groups = (
        final_df[
            [
                '_KOD_DUN',
                '_NAMA_DUN'
            ]
        ]
        .drop_duplicates()
        .copy()
    )

    dun_groups = dun_groups[
        dun_groups['_KOD_DUN'].ne('')
    ]

    dun_groups = dun_groups[
        dun_groups['_NAMA_DUN'].ne('')
    ]

    # --------------------------------------------------------
    # Sort DUN
    # --------------------------------------------------------

    dun_groups['_SORT'] = (
        dun_groups['_KOD_DUN']
        .apply(dun_sort_key)
    )

    dun_groups = (
        dun_groups
        .sort_values(
            '_SORT'
        )
        .drop(
            columns=['_SORT']
        )
    )

    # --------------------------------------------------------
    # Process every DUN
    # --------------------------------------------------------

    for (
        kod_dun,
        nama_dun
    ) in dun_groups.itertuples(
        index=False,
        name=None
    ):

        kod_dun = str(
            kod_dun
        ).strip()

        nama_dun = str(
            nama_dun
        ).strip()

        dun_df = final_df[
            (
                final_df['_KOD_DUN']
                == kod_dun
            )
            &
            (
                final_df['_NAMA_DUN']
                == nama_dun
            )
        ].copy()

        if dun_df.empty:
            continue

        # ----------------------------------------------------
        # DM GROUPS INSIDE DUN
        # ----------------------------------------------------

        dm_groups = (
            dun_df[
                [
                    '_KOD_DM',
                    '_NAMA_DM'
                ]
            ]
            .drop_duplicates()
            .copy()
        )

        dm_groups['_SORT'] = (
            dm_groups['_KOD_DM']
            .apply(dm_sort_key)
        )

        dm_groups = (
            dm_groups
            .sort_values(
                '_SORT'
            )
            .drop(
                columns=['_SORT']
            )
        )

        # ----------------------------------------------------
        # DM ROWS
        # ----------------------------------------------------

        dun_dm_rows = []

        for (
            kod_dm,
            nama_dm
        ) in dm_groups.itertuples(
            index=False,
            name=None
        ):

            kod_dm = str(
                kod_dm
            ).strip()

            nama_dm = str(
                nama_dm
            ).strip()

            dm_df = dun_df[
                (
                    dun_df['_KOD_DM']
                    == kod_dm
                )
                &
                (
                    dun_df['_NAMA_DM']
                    == nama_dm
                )
            ]

            if dm_df.empty:
                continue

            formatted_kod = format_kod_dm(
                kod_dm
            )

            row = build_stats_row(
                dm_df,
                kod=formatted_kod,
                nama=nama_dm
            )

            dun_dm_rows.append(row)

        # ----------------------------------------------------
        # Add DM rows
        # ----------------------------------------------------

        output_rows.extend(
            dun_dm_rows
        )

        # ----------------------------------------------------
        # DUN TOTAL
        # ----------------------------------------------------

        if dun_dm_rows:

            dun_dm_df = pd.DataFrame(
                dun_dm_rows
            )

            dun_total = build_total_row(
                dun_dm_df,
                label=nama_dun
            )

            # Use DUN code in KOD.
            dun_total['KOD'] = (
                kod_dun
            )

            output_rows.append(
                dun_total
            )

        # ----------------------------------------------------
        # Blank separator
        # ----------------------------------------------------

        output_rows.append(
            {
                header: ''
                for header in HEADERS
            }
        )

    # --------------------------------------------------------
    # GRAND TOTAL
    # --------------------------------------------------------

    all_dm_rows = [
        row
        for row in output_rows
        if row.get('JUMLAH', '') not in (
            '',
            None
        )
        and row.get('NAMA', '') not in (
            '',
            None
        )
    ]

    # Only use actual DM rows for grand total.
    actual_dm_rows = []

    for row in all_dm_rows:

        kod = str(
            row.get(
                'KOD',
                ''
            )
        )

        # A DM code has 3 components.
        if kod.count('/') == 2:

            actual_dm_rows.append(
                row
            )

    if actual_dm_rows:

        actual_dm_df = pd.DataFrame(
            actual_dm_rows
        )

        grand = build_grand_total(
            actual_dm_df
        )

        # Ensure it is clearly the grand total.
        grand['KOD'] = ''
        grand['NAMA'] = 'GRAND TOTAL'

        # Remove trailing separator before grand total.
        while (
            output_rows
            and output_rows[-1].get(
                'JUMLAH',
                ''
            ) in ('', None)
            and output_rows[-1].get(
                'NAMA',
                ''
            ) in ('', None)
        ):

            output_rows.pop()

        output_rows.append(
            grand
        )

    return pd.DataFrame(
        output_rows
    ).reindex(
        columns=HEADERS,
        fill_value=''
    )


# ============================================================
# PREPARE INPUT DATA
# ============================================================

def prepare_dataframe(
    df,
    fname,
    logs
):

    df = df.copy()

    # --------------------------------------------------------
    # Clean headers
    # --------------------------------------------------------

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    # --------------------------------------------------------
    # Find columns
    # --------------------------------------------------------

    col_nokp = get_col(
        df,
        [
            'nokp',
            'NoKp',
            'NO KP',
            'IC'
        ]
    )

    col_kod_parlimen = get_col(
        df,
        [
            'kod_parlimen',
            'KOD PARLIMEN',
            'KodParlimen'
        ]
    )

    col_nama_parlimen = get_col(
        df,
        [
            'nama_parlimen',
            'NAMA PARLIMEN',
            'NamaParlimen'
        ]
    )

    col_kod_dun = get_col(
        df,
        [
            'kod_dun',
            'KOD DUN',
            'KodDUN'
        ]
    )

    col_nama_dun = get_col(
        df,
        [
            'nama_dun',
            'NAMA DUN',
            'DUN',
            'NamaDUN'
        ]
    )

    col_kod_dm = get_col(
        df,
        [
            'kod_dm',
            'KOD DM',
            'KodDM'
        ]
    )

    col_nama_dm = get_col(
        df,
        [
            'nama_dm',
            'NAMA DM',
            'NamaDM'
        ]
    )

    col_jantina = get_col(
        df,
        [
            'jantina',
            'JANTINA'
        ]
    )

    col_race = get_col(
        df,
        [
            'kategori_kaum',
            'KATEGORI_KAUM',
            'kaum',
            'KAUM',
            'BANGSA'
        ]
    )

    col_umur = get_col(
        df,
        [
            'umur',
            'UMUR'
        ]
    )

    col_party = get_col(
        df,
        [
            'party',
            'PARTY',
            'parti'
        ]
    )

    col_sikap = get_col(
        df,
        [
            'sikap',
            'SIKAP',
            'catatan',
            'CATATAN'
        ]
    )

    col_service = get_col(
        df,
        [
            'noperkhidmatan',
            'NoPerkhidmatan',
            'NO PERKHIDMATAN'
        ]
    )

    col_pasangan = get_col(
        df,
        [
            'noperkhidmatanpasangan',
            'NoPerkhidmatanPasangan',
            'NoKPPasangan',
            'NOKP PASANGAN'
        ]
    )

    required = {
        'NoKp': col_nokp,
        'KOD PARLIMEN': col_kod_parlimen,
        'NAMA PARLIMEN': col_nama_parlimen,
        'KOD DUN': col_kod_dun,
        'NAMA DUN': col_nama_dun,
        'KOD DM': col_kod_dm,
        'NAMA DM': col_nama_dm,
        'JANTINA': col_jantina,
        'KAUM': col_race,
        'UMUR': col_umur,
        'PARTY': col_party,
        'SIKAP': col_sikap,
        'NoPerkhidmatan': col_service,
        'NoKPPasangan': col_pasangan
    }

    missing = [
        key
        for key, value in required.items()
        if value is None
    ]

    if missing:

        raise ValueError(
            f'{fname}: missing required columns: '
            + ', '.join(missing)
        )

    # --------------------------------------------------------
    # Internal columns
    # --------------------------------------------------------

    work = pd.DataFrame()

    work['_NOKP'] = (
        df[col_nokp]
        .fillna('')
        .astype(str)
        .str.strip()
    )

    work['_KOD_PARLIMEN'] = (
        df[col_kod_parlimen]
        .fillna('')
        .astype(str)
        .str.strip()
    )

    work['_NAMA_PARLIMEN'] = (
        df[col_nama_parlimen]
        .fillna('')
        .astype(str)
        .str.strip()
    )

    work['_KOD_DUN'] = (
        df[col_kod_dun]
        .fillna('')
        .astype(str)
        .str.strip()
    )

    work['_NAMA_DUN'] = (
        df[col_nama_dun]
        .fillna('')
        .astype(str)
        .str.strip()
    )

    work['_KOD_DM'] = (
        df[col_kod_dm]
        .fillna('')
        .astype(str)
        .str.strip()
    )

    work['_NAMA_DM'] = (
        df[col_nama_dm]
        .fillna('')
        .astype(str)
        .str.strip()
    )

    work['jantina_norm'] = (
        df[col_jantina]
        .fillna('')
        .astype(str)
        .str.strip()
        .str.upper()
    )

    work['race_norm'] = (
        df[col_race]
        .apply(normalise_race)
    )

    work['age_group'] = (
        df[col_umur]
        .apply(get_age_group)
    )

    work['party_norm'] = (
        df[col_party]
        .fillna('')
        .astype(str)
        .str.strip()
        .str.upper()
    )

    work['sikap_norm'] = (
        df[col_sikap]
        .apply(normalise_sikap)
    )

    # --------------------------------------------------------
    # Service numbers
    # --------------------------------------------------------

    work['NoPerkhidmatan_clean'] = (
        df[col_service]
        .apply(clean_service_no)
    )

    work['NoPerkhidmatanPasangan_clean'] = (
        df[col_pasangan]
        .apply(clean_service_no)
    )

    work['awal_type'] = (
        work['NoPerkhidmatan_clean']
        .apply(
            classify_pengundi_awal
        )
    )

    work['Pasangan Polis'] = (
        work[
            'NoPerkhidmatanPasangan_clean'
        ]
        .apply(
            lambda x:
                1
                if is_polis(x)
                else 0
        )
    )

    work['Pasangan Askar'] = (
        work[
            'NoPerkhidmatanPasangan_clean'
        ]
        .apply(
            lambda x:
                1
                if is_askar(x)
                else 0
        )
    )

    # --------------------------------------------------------
    # Clean invalid structural rows
    # --------------------------------------------------------

    work = work[
        work['_NOKP'].ne('')
    ].copy()

    work = work[
        work['_KOD_PARLIMEN'].ne('')
    ].copy()

    work = work[
        work['_KOD_DUN'].ne('')
    ].copy()

    work = work[
        work['_KOD_DM'].ne('')
    ].copy()

    logs.append(
        f'{fname}: {len(work):,} valid rows'
    )

    return work


# ============================================================
# MAIN GENERATOR
# ============================================================

def generate_demografik(uploaded_files):

    all_data = []
    logs = []

    # --------------------------------------------------------
    # Read every uploaded file
    # --------------------------------------------------------

    for uploaded_file in uploaded_files:

        fname = uploaded_file.name

        try:

            uploaded_file.seek(0)

            df = pd.read_excel(
                uploaded_file,
                dtype=str
            )

            if df.empty:

                logs.append(
                    f'{fname}: skipped — empty file'
                )

                continue

            prepared = prepare_dataframe(
                df,
                fname,
                logs
            )

            if not prepared.empty:

                all_data.append(
                    prepared
                )

        except Exception as e:

            logs.append(
                f'{fname}: ERROR — {e}'
            )

    # --------------------------------------------------------
    # No data
    # --------------------------------------------------------

    if not all_data:

        raise ValueError(
            'No valid data could be processed.\n\n'
            + '\n'.join(logs)
        )

    # --------------------------------------------------------
    # Combine files
    # --------------------------------------------------------

    final_df = pd.concat(
        all_data,
        ignore_index=True
    )

    logs.append(
        f'Total combined rows: '
        f'{len(final_df):,}'
    )

    # ========================================================
    # PARLIMEN
    # ========================================================

    df_parlimen = (
        build_parlimen_dataframe(
            final_df
        )
    )

    # ========================================================
    # DM GROUPED BY DUN
    # ========================================================

    df_dm = (
        build_dm_grouped_dataframe(
            final_df
        )
    )

    # ========================================================
    # CREATE WORKBOOK
    # ========================================================

    wb = Workbook()

    # Remove default worksheet.
    default_ws = wb.active

    wb.remove(
        default_ws
    )

    # --------------------------------------------------------
    # PARLIMEN
    # --------------------------------------------------------

    ws_parlimen = wb.create_sheet(
        title='PARLIMEN'
    )

    write_dataframe(
        ws_parlimen,
        df_parlimen
    )

    # --------------------------------------------------------
    # DM
    # --------------------------------------------------------

    ws_dm = wb.create_sheet(
        title='DM'
    )

    write_dataframe(
        ws_dm,
        df_dm
    )

    # ========================================================
    # ADD LIGHT TABLE STYLE
    # ========================================================
    #
    # Important:
    # TableStyleLight1 is used without custom header styling.
    # This keeps the workbook clean and avoids manually bolding
    # the headers.
    #
    # The DM sheet contains separator rows and DUN total rows,
    # so it is deliberately NOT converted into one Excel Table.
    #
    # ========================================================

    # PARLIMEN table only.
    if ws_parlimen.max_row >= 2:

        table_ref = (
            f'A1:AY{ws_parlimen.max_row}'
        )

        table = Table(
            displayName='ParlimenStats',
            ref=table_ref
        )

        style = TableStyleInfo(
            name='TableStyleLight1',
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=False,
            showColumnStripes=False
        )

        table.tableStyleInfo = style

        ws_parlimen.add_table(
            table
        )

    # ========================================================
    # FINAL CLEANUP
    # ========================================================

    for ws in wb.worksheets:

        for row in ws.iter_rows():

            for cell in row:

                cell.font = Font(
                    name='Calibri',
                    size=11,
                    bold=False
                )

    # ========================================================
    # SAVE TO MEMORY
    # ========================================================

    output = io.BytesIO()

    wb.save(
        output
    )

    output.seek(0)

    # ========================================================
    # OUTPUT NAME
    # ========================================================

    # Try to use first DUN name if available.
    dun_names = (
        final_df['_NAMA_DUN']
        .dropna()
        .astype(str)
        .str.strip()
    )

    dun_names = [
        x
        for x in dun_names.unique()
        if x
    ]

    if len(dun_names) == 1:

        out_name = (
            f'DEMOGRAFIK '
            f'{clean_filename(dun_names[0])}.xlsx'
        )

    else:

        out_name = (
            f'DEMOGRAFIK '
            f'({len(dun_names)} DUN).xlsx'
        )

    logs.append(
        'Workbook created with sheets: '
        'PARLIMEN, DM'
    )

    logs.append(
        'DUN worksheet removed.'
    )

    logs.append(
        'DM worksheet grouped by DUN with '
        'DUN totals and separator rows.'
    )

    return (
        output.getvalue(),
        out_name,
        logs
    )
