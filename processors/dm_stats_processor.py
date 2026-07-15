import io
import re
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo

st.set_page_config(page_title="Demografik Generator", layout="wide")

st.title("DEMOGRAFIK Generator")

uploaded_files = st.file_uploader(
    "Upload Excel file(s)",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

HEADERS = [
    'KOD DM', 'NAMA DM', 'JUMLAH',
    'LELAKI', 'LELAKI (%)', 'PEREMPUAN', 'PEREMPUAN (%)',
    'MELAYU', 'MELAYU (%)', 'CINA', 'CINA (%)', 'INDIA', 'INDIA (%)', 'LAIN-LAIN', 'LAIN-LAIN (%)',
    '18-24', '18-24 (%)', '25-30', '25-30 (%)', '31-40', '31-40 (%)',
    '41-50', '41-50 (%)', '51-60', '51-60 (%)', '61+', '61+ (%)',
    'PAS', 'PAS (%)', 'PKR', 'PKR (%)', 'PPBM', 'PPBM (%)', 'UMNO', 'UMNO (%)',
    'PUTIH', 'PUTIH (%)', 'KELABU', 'KELABU (%)', 'HITAM', 'HITAM (%)',
    'PENGUNDI AWAL', 'PENGUNDI AWAL (%)',
    'POLIS', 'POLIS (%)',
    'PASANGAN POLIS', 'PASANGAN POLIS (%)',
    'ASKAR', 'ASKAR (%)',
    'PASANGAN ASKAR', 'PASANGAN ASKAR (%)'
]

MAIN_RACES = ['MELAYU', 'CINA', 'INDIA', 'LAIN-LAIN']
AGE_GROUPS = ['18-21', '22-30', '31-40', '41-50', '51-60', '61+']
PARTY_COLS = ['PAS', 'PKR', 'PPBM', 'UMNO']
SIKAP_COLS = ['PUTIH', 'KELABU', 'HITAM']


def get_col(df, possible_names):
    col_map = {c.lower().strip(): c for c in df.columns}
    for name in possible_names:
        key = name.lower().strip()
        if key in col_map:
            return col_map[key]
    return None


def clean_service_no(value):
    n = str(value).strip().upper()
    return '' if n in {'', 'NAN', 'NONE', 'NULL'} else n


def clean_filename(value):
    name = str(value).strip().upper()
    name = re.sub(r'[\\/:*?"<>|]', ' ', name)
    name = ' '.join(name.split())
    return name if name else 'OUTPUT'


def clean_sheet_name(value, existing_names):
    """Sanitize a worksheet name (Excel: max 31 chars, no \\/?*[]:), and
    make sure it's unique within the workbook."""
    name = str(value).strip().upper()
    name = re.sub(r'[\\/?*\[\]:]', ' ', name)
    name = ' '.join(name.split())
    if not name:
        name = 'DUN'
    name = name[:31]

    base = name
    counter = 2
    while name in existing_names:
        suffix = f" ({counter})"
        name = base[:31 - len(suffix)] + suffix
        counter += 1

    existing_names.add(name)
    return name


def kod_dun_digits(kod_dun):
    """Strip a kod_dun value down to just its digits (handles floats like
    '14.0' from Excel and any stray non-numeric characters)."""
    kod_str = str(kod_dun).strip().split('.')[0]
    return re.sub(r'\D', '', kod_str)


def kod_dun_sort_key(kod_dun):
    digits = kod_dun_digits(kod_dun)
    return int(digits) if digits else -1


def format_sheet_label(kod_dun, nama_dun):
    digits = kod_dun_digits(kod_dun)
    last2 = digits[-2:].zfill(2) if digits else '00'
    return f"N.{last2} {nama_dun}"


def format_kod_dm(value):
    kod = str(value).strip()
    if kod in {'', 'None', 'nan', 'NaN'}:
        return ''
    kod = kod.split('.')[0].zfill(7)
    return f"{kod[:3]}/{kod[3:5]}/{kod[5:]}"


def normalise_race(value):
    r = str(value).strip().upper()
    return r if r in {'MELAYU', 'CINA', 'INDIA'} else 'LAIN-LAIN'


def normalise_sikap(value):
    s = str(value).strip().upper()
    if s in {'KELABU-LAMA', 'KELABU-BARU'}:
        return 'KELABU'
    if s in {'PUTIH', 'KELABU', 'HITAM'}:
        return s
    return ''


def classify_awal(value):
    n = clean_service_no(value)
    if n == '':
        return ''
    if n.startswith('G') or n.startswith('RF'):
        return 'POLIS'
    if n.startswith('T'):
        return 'ASKAR'
    return 'PENGUNDI AWAL'


def is_polis(value):
    n = clean_service_no(value)
    return n.startswith('G') or n.startswith('RF')


def is_askar(value):
    n = clean_service_no(value)
    return n.startswith('T')


def get_age_group(value):
    try:
        a = int(float(value))
        if 18 <= a <= 21:
            return '18-21'
        elif 22 <= a <= 30:
            return '22-30'
        elif 31 <= a <= 40:
            return '31-40'
        elif 41 <= a <= 50:
            return '41-50'
        elif 51 <= a <= 60:
            return '51-60'
        elif a >= 61:
            return '61+'
    except Exception:
        pass
    return ''


def pct(part, total):
    return round(part / total * 100, 1) if total else 0


def build_dm_row(kod_dm, nama_dm, grp):
    total = len(grp)

    row = {
        'KOD DM': format_kod_dm(kod_dm),
        'NAMA DM': nama_dm,
        'JUMLAH': total
    }

    sex_vc = grp['_jantina'].value_counts()
    race_vc = grp['_race'].value_counts()
    age_vc = grp['_age_group'].value_counts()
    party_vc = grp['_party'].value_counts()
    sikap_vc = grp['_sikap'].value_counts()
    awal_vc = grp['_awal_type'].value_counts()

    for key, label in [('L', 'LELAKI'), ('P', 'PEREMPUAN')]:
        c = sex_vc.get(key, 0)
        row[label] = c
        row[f'{label} (%)'] = pct(c, total)

    for r in MAIN_RACES:
        c = race_vc.get(r, 0)
        row[r] = c
        row[f'{r} (%)'] = pct(c, total)

    for a in AGE_GROUPS:
        c = age_vc.get(a, 0)
        row[a] = c
        row[f'{a} (%)'] = pct(c, total)

    for p in PARTY_COLS:
        c = party_vc.get(p, 0)
        row[p] = c
        row[f'{p} (%)'] = pct(c, total)

    for s in SIKAP_COLS:
        c = sikap_vc.get(s, 0)
        row[s] = c
        row[f'{s} (%)'] = pct(c, total)

    pengundi_awal = grp['_NoPerkhidmatan_clean'].ne('').sum()
    polis = awal_vc.get('POLIS', 0)
    askar = awal_vc.get('ASKAR', 0)

    pasangan_polis = grp['_Pasangan Polis'].sum()
    pasangan_askar = grp['_Pasangan Askar'].sum()

    row['PENGUNDI AWAL'] = pengundi_awal
    row['PENGUNDI AWAL (%)'] = pct(pengundi_awal, total)

    row['POLIS'] = polis
    row['POLIS (%)'] = pct(polis, total)

    row['PASANGAN POLIS'] = pasangan_polis
    row['PASANGAN POLIS (%)'] = pct(pasangan_polis, total)

    row['ASKAR'] = askar
    row['ASKAR (%)'] = pct(askar, total)

    row['PASANGAN ASKAR'] = pasangan_askar
    row['PASANGAN ASKAR (%)'] = pct(pasangan_askar, total)

    return row


def add_total_row(df):
    total = df['JUMLAH'].sum()
    total_row = {'KOD DM': '', 'NAMA DM': '', 'JUMLAH': total}

    for h in HEADERS:
        if h in ['KOD DM', 'NAMA DM', 'JUMLAH']:
            continue

        if h.endswith('(%)'):
            base = h.replace(' (%)', '')
            total_row[h] = pct(total_row.get(base, 0), total)
        else:
            total_row[h] = pd.to_numeric(df[h], errors='coerce').fillna(0).sum()

    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def build_rumusan_df(dun_df):
    """Build the summary (one row per DM) table for a single DUN's data."""
    rows = []
    for (kod_dm, nama_dm), grp in dun_df.groupby(['_KOD DM', '_NAMA DM'], dropna=False):
        rows.append(build_dm_row(kod_dm, nama_dm, grp))

    rumusan_df = pd.DataFrame(rows)

    for h in HEADERS:
        if h not in rumusan_df.columns:
            rumusan_df[h] = 0

    rumusan_df = rumusan_df[HEADERS]
    rumusan_df = rumusan_df.sort_values(by='KOD DM', kind='stable')
    rumusan_df = add_total_row(rumusan_df)
    return rumusan_df


def write_dun_sheet(ws, rumusan_df):
    """Write and style a single DUN's rumusan_df onto the given worksheet."""
    BLUE = '9DC3E6'
    GREEN = 'A9D18E'
    ORANGE = 'F4B183'
    YELLOW = 'FFD966'
    PURPLE = 'B4A7D6'
    WHITE = 'D9D9D9'

    thin = Side(style='thin', color='000000')
    medium = Side(style='medium', color='000000')

    group_fill = {}
    for c in range(4, 8):
        group_fill[c] = BLUE
    for c in range(8, 16):
        group_fill[c] = GREEN
    for c in range(16, 28):
        group_fill[c] = ORANGE
    for c in range(28, 36):
        group_fill[c] = YELLOW
    for c in range(36, 42):
        group_fill[c] = WHITE
    for c in range(42, 52):
        group_fill[c] = PURPLE

    group_left_edges = {1, 4, 8, 16, 28, 36, 42, 44, 46, 48, 50}
    group_right_edges = {2, 7, 15, 27, 35, 41, 43, 45, 47, 49, 51}
    thin_right_edges = {5, 9, 11, 13, 17, 19, 21, 23, 25, 29, 31, 33, 37, 39}
    thin_left_edges = {10, 12, 14, 18, 20, 22, 24, 26, 30, 32, 34, 38, 40}

    for col_idx, h in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = Font(name='Calibri', size=11, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        if col_idx in group_fill:
            cell.fill = PatternFill('solid', fgColor=group_fill[col_idx])

        left = medium if col_idx in group_left_edges else (thin if col_idx in thin_left_edges else thin)
        right = medium if col_idx in group_right_edges else (thin if col_idx in thin_right_edges else thin)

        cell.border = Border(left=left, right=right, top=medium, bottom=medium)

    for r_idx, row in enumerate(rumusan_df.itertuples(index=False), start=2):
        is_total = r_idx == len(rumusan_df) + 1

        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)

            if c_idx == 2:
                cell.alignment = Alignment(horizontal='left', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')

            cell.font = Font(name='Calibri', size=11, bold=is_total)

            if c_idx in group_fill and is_total:
                cell.fill = PatternFill('solid', fgColor=group_fill[c_idx])

            left = medium if c_idx in group_left_edges else (thin if c_idx in thin_left_edges else thin)
            right = medium if c_idx in group_right_edges else (thin if c_idx in thin_right_edges else thin)

            cell.border = Border(
                left=left,
                right=right,
                top=medium if is_total else thin,
                bottom=medium if is_total else thin
            )

            col_name = HEADERS[c_idx - 1]

            if isinstance(value, (int, float)):
                if '(%)' in col_name:
                    cell.number_format = '0.0'
                else:
                    cell.number_format = '#,##0'

    widths = {
        'A': 13, 'B': 25, 'C': 13,
        'D': 11.3, 'E': 14.7, 'F': 17, 'G': 20.6,
        'H': 13.1, 'I': 16.6, 'J': 10, 'K': 13.4, 'L': 10.7, 'M': 14.1, 'N': 14.6, 'O': 18.1,
        'P': 10.3, 'Q': 13.7, 'R': 10.3, 'S': 13.7, 'T': 10.3, 'U': 13.7,
        'V': 10.3, 'W': 13.7, 'X': 10.3, 'Y': 13.7, 'Z': 8.6, 'AA': 12,
        'AB': 9, 'AC': 12.4, 'AD': 9, 'AE': 12.4, 'AF': 10.9, 'AG': 14.3, 'AH': 11.7, 'AI': 15.1,
        'AJ': 11, 'AK': 14.4, 'AL': 12.4, 'AM': 15.9, 'AN': 11.6, 'AO': 15,
        'AP': 21.3, 'AQ': 24.9, 'AR': 10.6, 'AS': 14, 'AT': 21.4, 'AU': 25,
        'AV': 11.4, 'AW': 14.9, 'AX': 22.4, 'AY': 26
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    max_len = 0
    for row in range(1, ws.max_row + 1):
        value = ws.cell(row=row, column=2).value
        max_len = max(max_len, len(str(value or '')))
    ws.column_dimensions['B'].width = min(max_len + 4, 60)

    for row in range(2, ws.max_row + 1):
        ws.cell(row=row, column=2).alignment = Alignment(horizontal='left', vertical='center')

    ws.row_dimensions[1].height = 15.75
    ws.freeze_panes = 'A2'

    end_row = len(rumusan_df) + 1
    table_ref = f"A1:AY{end_row}"

    # Table display names must also be unique within the workbook.
    tab = Table(displayName=f"Table_{ws.title.replace(' ', '_')[:20]}_{ws.parent.index(ws)}", ref=table_ref)
    style = TableStyleInfo(
        name="TableStyleLight1",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=False,
        showColumnStripes=False
    )
    tab.tableStyleInfo = style
    ws.add_table(tab)


def generate_demografik(uploaded_files):
    all_data = []
    logs = []

    for uploaded_file in uploaded_files:
        fname = uploaded_file.name

        try:
            df = pd.read_excel(uploaded_file, dtype=str)
            df.columns = [c.strip() for c in df.columns]

            col_dm = get_col(df, ['KOD DM', 'kod_dm'])
            col_nama_dm = get_col(df, ['NamaDM', 'nama_dm', 'NAMA DM'])
            col_kod_dun = get_col(df, ['kod_dun', 'KOD DUN', 'KODDUN'])
            col_nama_dun = get_col(df, ['nama_dun', 'DUN', 'NAMA DUN'])
            col_jantina = get_col(df, ['JANTINA', 'jantina'])
            col_bangsa = get_col(df, ['kaum', 'BANGSA', 'kategori_kaum'])
            col_umur = get_col(df, ['UMUR', 'umur'])
            col_party = get_col(df, ['party', 'PARTY'])
            col_sikap = get_col(df, ['CATATAN', 'sikap'])
            col_no = get_col(df, ['NoPerkhidmatan', 'noperkhidmatan'])
            col_pasangan = get_col(df, ['NoKPPasangan', 'NoPerkhidmatanPasangan', 'noperkhidmatanpasangan'])

            required = {
                'KOD DM': col_dm,
                'NamaDM': col_nama_dm,
                'kod_dun': col_kod_dun,
                'nama_dun / DUN': col_nama_dun,
                'JANTINA': col_jantina,
                'BANGSA': col_bangsa,
                'UMUR': col_umur,
                'NoPerkhidmatan': col_no,
                'NoKPPasangan': col_pasangan
            }

            missing = [k for k, v in required.items() if v is None]
            if missing:
                logs.append(f"Skipped {fname} — missing columns: {missing}")
                continue

            df['_KOD_DUN'] = df[col_kod_dun].fillna('').astype(str).str.strip()
            df['_NAMA_DUN'] = df[col_nama_dun].fillna('').astype(str).str.strip().str.upper()
            df['_KOD DM'] = df[col_dm].fillna('').astype(str).str.strip()
            df['_NAMA DM'] = df[col_nama_dm].fillna('').astype(str).str.strip()
            df['_jantina'] = df[col_jantina].fillna('').astype(str).str.strip().str.upper()
            df['_race'] = df[col_bangsa].apply(normalise_race)
            df['_age_group'] = df[col_umur].apply(get_age_group)

            df['_party'] = df[col_party].fillna('').astype(str).str.strip().str.upper() if col_party else ''
            df['_sikap'] = df[col_sikap].apply(normalise_sikap) if col_sikap else ''

            df['_NoPerkhidmatan_clean'] = df[col_no].apply(clean_service_no)
            df['_NoKPPasangan_clean'] = df[col_pasangan].apply(clean_service_no)

            df['_awal_type'] = df['_NoPerkhidmatan_clean'].apply(classify_awal)
            df['_Pasangan Polis'] = df['_NoKPPasangan_clean'].apply(lambda x: 1 if is_polis(x) else 0)
            df['_Pasangan Askar'] = df['_NoKPPasangan_clean'].apply(lambda x: 1 if is_askar(x) else 0)

            all_data.append(df)
            logs.append(f"Loaded {fname}: {len(df):,} rows")

        except Exception as e:
            logs.append(f"Error reading {fname}: {e}")

    if not all_data:
        raise ValueError("No valid data loaded.\n" + "\n".join(logs))

    final_df = pd.concat(all_data, ignore_index=True)

    # Unique (kod_dun, nama_dun) combos, sorted by kod_dun ascending.
    dun_combos = (
        final_df[['_KOD_DUN', '_NAMA_DUN']]
        .drop_duplicates()
    )
    dun_combos = dun_combos[(dun_combos['_KOD_DUN'] != '') & (dun_combos['_NAMA_DUN'] != '')]
    dun_combos = sorted(
        dun_combos.itertuples(index=False, name=None),
        key=lambda pair: kod_dun_sort_key(pair[0])
    )

    if not dun_combos:
        raise ValueError("No DUN name/kod_dun found in the uploaded data.")

    wb = Workbook()
    wb.remove(wb.active)  # drop the default blank sheet

    existing_sheet_names = set()

    for kod_dun, nama_dun in dun_combos:
        dun_df = final_df[(final_df['_KOD_DUN'] == kod_dun) & (final_df['_NAMA_DUN'] == nama_dun)]
        rumusan_df = build_rumusan_df(dun_df)

        sheet_label = format_sheet_label(kod_dun, nama_dun)
        sheet_name = clean_sheet_name(sheet_label, existing_sheet_names)
        ws = wb.create_sheet(title=sheet_name)
        write_dun_sheet(ws, rumusan_df)

        logs.append(f"Built sheet '{sheet_name}': {len(dun_df):,} rows, {len(rumusan_df) - 1} DM(s)")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    if len(dun_combos) == 1:
        kod_dun, nama_dun = dun_combos[0]
        out_name = f"DEMOGRAFIK {clean_filename(format_sheet_label(kod_dun, nama_dun))}.xlsx"
    else:
        out_name = f"DEMOGRAFIK ({len(dun_combos)} DUN).xlsx"

    return output.getvalue(), out_name, logs


if uploaded_files:
    if st.button("Generate DEMOGRAFIK"):
        try:
            excel_bytes, out_name, logs = generate_demografik(uploaded_files)

            st.success(f"Generated: {out_name}")

            with st.expander("Processing log"):
                for log in logs:
                    st.write(log)

            st.download_button(
                label="Download Excel",
                data=excel_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(str(e))
else:
    st.info("Upload one or more Excel files to start.")
