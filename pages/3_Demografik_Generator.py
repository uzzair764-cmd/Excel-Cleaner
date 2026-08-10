import io
import openpyxl
import streamlit as st
import processors.dm_stats_processor as dm_stats
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="DM Stats", page_icon="📊", layout="wide")
st.title("📊 DEMOGRAFIK Generator")

uploaded_files = st.file_uploader("Upload Excel file(s)", type=["xlsx", "xls"], accept_multiple_files=True, key="dm_stats_excel_uploader")

st.subheader("Age Group Configuration")
age_group_text = st.text_input("Age groups", value="18-21, 22-30, 31-40, 41-50, 51-60, 61+", help="Enter age groups separated by commas. Example: 18-24, 25-30, 31-40, 41-50, 51-60, 61+")
age_groups = dm_stats.parse_age_groups(age_group_text)

if age_groups:
    st.caption("Active age groups: " + " | ".join(age_groups))
else:
    st.error("Invalid age group configuration. Example: 18-21, 22-30, 31-40, 41-50, 51-60, 61+")


# ============================================================
# AGE GROUP × KAUM × JANTINA TABLE
# ============================================================
# Each race block is now:
#
# JUMLAH | 18-21 | 18-21 (%) | 22-30 | 22-30 (%) | ...
# Lelaki | count  | %         | ...
# Perempuan | ...
# MELAYU | count  | %         | ...
#
# The same block repeats for CINA, INDIA and LAIN-LAIN.
# Only the header rows are highlighted.
# Tables start at column B and have thick outside borders.
# ============================================================


def _build_age_race_gender_rows_updated(data_df, age_groups):
    rows = []

    for race in dm_stats.MAIN_RACES:
        race_df = data_df[data_df['_race'] == race]

        header = {'LABEL': 'JUMLAH'}
        for age in age_groups:
            header[age] = age
            header[f'{age} (%)'] = f'{age} (%)'
        rows.append(('header', header))

        for gender_code, label in [('L', 'Lelaki'), ('P', 'Perempuan')]:
            gender_df = race_df[race_df['_jantina'] == gender_code]
            total = len(gender_df)
            data = {'LABEL': label, 'JUMLAH': total}
            for age in age_groups:
                count = int((gender_df['_age_group'] == age).sum())
                data[age] = count
                data[f'{age} (%)'] = dm_stats.pct(count, total)
            rows.append(('gender', data))

        race_total = len(race_df)
        total_data = {'LABEL': race, 'JUMLAH': race_total}
        for age in age_groups:
            count = int((race_df['_age_group'] == age).sum())
            total_data[age] = count
            total_data[f'{age} (%)'] = dm_stats.pct(count, race_total)
        rows.append(('race_total', total_data))

    return rows


dm_stats.build_age_race_gender_rows = _build_age_race_gender_rows_updated


def _write_age_table_formatted(ws, start_row, title_text, data_df, age_groups):
    thin = Side(style='thin', color='000000')
    medium = Side(style='medium', color='000000')
    GREEN = 'A9D18E'
    ORANGE = 'F4B183'
    WHITE = 'FFFFFF'

    start_col = 2
    end_col = start_col + len(age_groups) * 2

    title = ws.cell(start_row, start_col, title_text)
    ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=end_col)
    title.font = Font(name='Calibri', size=14, bold=True)
    title.alignment = Alignment(horizontal='center', vertical='center')
    title.fill = PatternFill(fill_type=None)
    ws.row_dimensions[start_row].height = 24

    for c in range(start_col, end_col + 1):
        ws.cell(start_row, c).border = Border(
            left=medium if c == start_col else thin,
            right=medium if c == end_col else thin,
            top=medium,
            bottom=medium
        )

    current = start_row + 1
    table_first_row = current

    for row_type, data in dm_stats.build_age_race_gender_rows(data_df, age_groups):
        values = [data['LABEL']]
        for age in age_groups:
            values.extend([data[age], data[f'{age} (%)']])

        for offset, value in enumerate(values):
            c = start_col + offset
            cell = ws.cell(current, c, value)
            cell.font = Font(name='Calibri', size=11, bold=(row_type == 'header'))
            cell.alignment = Alignment(horizontal='left' if offset == 0 else 'center', vertical='center', wrap_text=True)

            if row_type == 'header':
                cell.fill = PatternFill(fill_type='solid', fgColor=GREEN if offset == 0 else ORANGE)
            else:
                cell.fill = PatternFill(fill_type='solid', fgColor=WHITE)

            cell.border = Border(
                left=medium if c == start_col else thin,
                right=medium if c == end_col else thin,
                top=thin,
                bottom=thin
            )

            if row_type in {'gender', 'race_total'} and offset > 0 and isinstance(value, (int, float)):
                cell.number_format = '#,##0' if offset % 2 == 1 else '0.0"%"'

        current += 1

    table_last_row = current - 1

    for r in range(table_first_row, table_last_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(r, c)
            cell.border = Border(
                left=medium if c == start_col else thin,
                right=medium if c == end_col else thin,
                top=medium if r == table_first_row else thin,
                bottom=medium if r == table_last_row else thin
            )

    ws.column_dimensions['B'].width = max(ws.column_dimensions['B'].width or 0, 18)
    for i in range(len(age_groups)):
        count_col = get_column_letter(start_col + 1 + i * 2)
        pct_col = get_column_letter(start_col + 2 + i * 2)
        ws.column_dimensions[count_col].width = max(ws.column_dimensions[count_col].width or 0, 11)
        ws.column_dimensions[pct_col].width = max(ws.column_dimensions[pct_col].width or 0, 13.7)

    return current + 1


def _write_dun_age_table_formatted(ws, start_row, kod_dun, nama_dun, dun_df, age_groups):
    title = f"N.{dm_stats.kod_dun_digits(kod_dun)[-2:].zfill(2)} {str(nama_dun).upper()} - UMUR MENGIKUT KAUM DAN JANTINA"
    return _write_age_table_formatted(ws, start_row, title, dun_df, age_groups)


def _write_parliament_age_table_formatted(ws, start_row, kod_parlimen, nama_parlimen, parliament_df, age_groups):
    title = f"{str(kod_parlimen).strip()} {str(nama_parlimen).upper()} - UMUR MENGIKUT KAUM DAN JANTINA"
    return _write_age_table_formatted(ws, start_row, title, parliament_df, age_groups)


dm_stats.write_dun_age_table = _write_dun_age_table_formatted
dm_stats.write_parliament_age_table = _write_parliament_age_table_formatted


def apply_number_formatting(excel_bytes):
    input_buffer = io.BytesIO(excel_bytes)
    wb = openpyxl.load_workbook(input_buffer)

    for ws in wb.worksheets:
        header_map = {col: ws.cell(1, col).value for col in range(1, ws.max_column + 1)}
        for col, header in header_map.items():
            if header is None:
                continue
            is_percentage = '(%)' in str(header)
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row, col)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '0.0"%"' if is_percentage else '#,##0'

        for merged_range in list(ws.merged_cells.ranges):
            if merged_range.min_col != 2 or merged_range.max_row != merged_range.min_row:
                continue
            title_cell = ws.cell(merged_range.min_row, 2)
            if not isinstance(title_cell.value, str) or 'UMUR MENGIKUT KAUM DAN JANTINA' not in title_cell.value.upper():
                continue

            title_row = merged_range.min_row
            start_col = 2
            end_col = merged_range.max_col
            table_start = title_row + 1
            table_end = ws.max_row

            for next_range in list(ws.merged_cells.ranges):
                if (
                    next_range.min_col == 2
                    and next_range.min_row > title_row
                    and next_range.min_row < table_end
                    and isinstance(ws.cell(next_range.min_row, 2).value, str)
                    and 'UMUR MENGIKUT KAUM DAN JANTINA' in ws.cell(next_range.min_row, 2).value.upper()
                ):
                    table_end = next_range.min_row - 2
                    break

            for r in range(table_start, table_end + 1):
                for c in range(start_col, end_col + 1):
                    cell = ws.cell(r, c)
                    if isinstance(cell.value, (int, float)) and c > start_col:
                        offset = c - start_col
                        cell.number_format = '0.0"%"' if offset % 2 == 0 else '#,##0'
                    cell.border = Border(
                        left=Side(style='medium', color='000000') if c == start_col else Side(style='thin', color='000000'),
                        right=Side(style='medium', color='000000') if c == end_col else Side(style='thin', color='000000'),
                        top=Side(style='medium', color='000000') if r == table_start else Side(style='thin', color='000000'),
                        bottom=Side(style='medium', color='000000') if r == table_end else Side(style='thin', color='000000')
                    )

            for c in range(start_col, end_col + 1):
                ws.cell(title_row, c).border = Border(
                    left=Side(style='medium', color='000000') if c == start_col else Side(style='thin', color='000000'),
                    right=Side(style='medium', color='000000') if c == end_col else Side(style='thin', color='000000'),
                    top=Side(style='medium', color='000000'),
                    bottom=Side(style='medium', color='000000')
                )

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer.getvalue()


if uploaded_files and age_groups:
    if st.button("Generate DEMOGRAFIK", key="dm_stats_generate_button"):
        try:
            excel_bytes, out_name, logs = dm_stats.generate_demografik(uploaded_files, age_groups=age_groups)
            excel_bytes = apply_number_formatting(excel_bytes)
            st.success(f"Generated: {out_name}")
            with st.expander("Processing log"):
                for log in logs:
                    st.write(log)
            st.download_button(
                label="Download Excel",
                data=excel_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dm_stats_download_button"
            )
        except Exception as e:
            st.error(str(e))
elif not uploaded_files:
    st.info("Upload one or more Excel files to start.")
