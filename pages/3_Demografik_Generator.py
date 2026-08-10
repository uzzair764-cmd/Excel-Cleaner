import io
import pandas as pd


RACES = ["MELAYU", "CINA", "INDIA", "LAIN-LAIN"]
AGE_GROUPS = ["18-25", "26-40", "41-60", "61>"]
GENDERS = ["L", "P"]
TEL_COLS = ["PEMILIH", "POLIS", "ASKAR"]


def normalise_race(value):
    r = str(value).strip().upper()
    if r in ["MELAYU", "CINA", "INDIA"]:
        return r
    return "LAIN-LAIN"


def age_group(value):
    try:
        a = int(float(value))
        if 18 <= a <= 25:
            return "18-25"
        if 26 <= a <= 40:
            return "26-40"
        if 41 <= a <= 60:
            return "41-60"
        if a >= 61:
            return "61>"
    except Exception:
        pass
    return ""


def pct(part, total):
    return part / total if total else 0


def get_service_col(df):
    for col in df.columns:
        if str(col).strip().lower() == "noperkhidmatan":
            return col
    return None


def prepare_demografik_df(raw_df, cleaned_df):
    raw = raw_df.copy()
    cleaned = cleaned_df.copy()

    needed_cols = [
        "nokp", "umur", "jantina", "kaum_spr",
        "party", "number", "kod_parlimen", "nama_parlimen",
        "kod_dun", "nama_dun", "kod_dm", "nama_dm"
    ]

    for col in needed_cols:
        if col not in raw.columns:
            raw[col] = ""

    for col in needed_cols:
        if col not in cleaned.columns:
            cleaned[col] = ""

    raw["_race"] = raw["kaum_spr"].apply(normalise_race)
    raw["_age"] = raw["umur"].apply(age_group)
    raw["_jantina"] = raw["jantina"].astype(str).str.strip().str.upper()

    cleaned["_tel_count"] = 1

    service_col = get_service_col(cleaned)
    
    if service_col:
        svc = cleaned[service_col].astype(str).str.strip().str.upper()
    else:
        svc = pd.Series([""] * len(cleaned), index=cleaned.index)

    nokp_svc = cleaned["nokp"].astype(str).str.strip().str.upper()

    cleaned["_polis"] = (
        svc.str.startswith(("G", "R"), na=False) |
        nokp_svc.str.startswith(("G", "R"), na=False)
    )

    cleaned["_askar"] = (
        svc.str.startswith("T", na=False) |
        nokp_svc.str.startswith("T", na=False)
    )

    tel_summary = (
        cleaned
        .groupby(
            ["kod_parlimen", "nama_parlimen", "kod_dun", "nama_dun", "kod_dm", "nama_dm"],
            dropna=False
        )
        .agg(
            _pemilih_tel=("_tel_count", "sum"),
            _polis=("_polis", "sum"),
            _askar=("_askar", "sum")
        )
        .reset_index()
    )

    return raw, tel_summary


def get_tel_counts_for_group(tel_summary, keys):
    if tel_summary.empty:
        return 0, 0, 0

    mask = pd.Series([True] * len(tel_summary), index=tel_summary.index)

    for col, val in keys.items():
        mask &= tel_summary[col].astype(str).fillna("").eq(str(val))

    subset = tel_summary.loc[mask]

    if subset.empty:
        return 0, 0, 0

    return (
        int(pd.to_numeric(subset["_pemilih_tel"], errors="coerce").fillna(0).sum()),
        int(pd.to_numeric(subset["_polis"], errors="coerce").fillna(0).sum()),
        int(pd.to_numeric(subset["_askar"], errors="coerce").fillna(0).sum())
    )


def build_demo_row(code, name, group, tel_counts):
    total = len(group)

    race_vc = group["_race"].value_counts()
    age_vc = group["_age"].value_counts()
    gender_vc = group["_jantina"].value_counts()

    pemilih, polis, askar = tel_counts

    row = {
        "KOD": code,
        "NAMA": name,
        "JUMLAH PENGUNDI": total,
    }

    for race in RACES:
        row[race] = int(race_vc.get(race, 0))

    for age in AGE_GROUPS:
        row[age] = int(age_vc.get(age, 0))

    for gender in GENDERS:
        row[gender] = int(gender_vc.get(gender, 0))

    row["PEMILIH"] = pemilih
    row["POLIS"] = polis
    row["ASKAR"] = askar

    return row


def split_total_and_percent(rows):
    df = pd.DataFrame(rows)

    if df.empty:
        return df, {}, {}

    count_cols = ["JUMLAH PENGUNDI"] + RACES + AGE_GROUPS + GENDERS + TEL_COLS

    total_row = {"LABEL": "JUMLAH PENGUNDI"}
    percent_row = {"LABEL": "PERATUSAN"}

    for col in count_cols:
        total_row[col] = int(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())

    grand_total = total_row["JUMLAH PENGUNDI"]

    percent_row["JUMLAH PENGUNDI"] = ""

    for col in RACES + AGE_GROUPS + GENDERS + TEL_COLS:
        percent_row[col] = pct(total_row[col], grand_total)

    return df, total_row, percent_row


def format_parlimen_label(df, base_name):
    kod = ""
    nama = ""

    kod_series = df["kod_parlimen"].replace("", pd.NA).dropna()
    nama_series = df["nama_parlimen"].replace("", pd.NA).dropna()

    if not kod_series.empty:
        kod = str(kod_series.iloc[0]).strip().zfill(3)

    if not nama_series.empty:
        nama = str(nama_series.iloc[0]).strip().upper()

    if kod and nama:
        return f"P.{kod} {nama}"

    return base_name.upper()


def format_dun_code(kod_dun):
    kod = str(kod_dun).strip()
    if not kod:
        return ""
    return f"N.{kod[-2:].zfill(2)}"


def format_dm_code(kod_dm):
    kod = str(kod_dm).strip()
    if not kod:
        return ""
    return f"DM.{kod[-2:].zfill(2)}"


def make_formats(workbook):
    return {
        "title": workbook.add_format({
            "font_name": "Calibri",
            "font_size": 22,
            "bold": True,
            "font_color": "#000000",
            "align": "left",
            "valign": "vcenter"
        }),
        "subtitle": workbook.add_format({
            "font_name": "Calibri",
            "font_size": 22,
            "bold": True,
            "font_color": "#000000",
            "align": "left",
            "valign": "vcenter"
        }),
        "legend_melayu": workbook.add_format({
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": "#000000",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#FFFFFF"
        }),
        "legend_bukan": workbook.add_format({
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": "#000000",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#F4B084"
        }),
        "header": workbook.add_format({
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": "#000000",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#D9D9D9"
        }),
        "center": workbook.add_format({
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": "#000000",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "num_format": "#,##0"
        }),
        "left": workbook.add_format({
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": "#000000",
            "border": 1,
            "align": "left",
            "valign": "vcenter",
            "num_format": "#,##0"
        }),
        "orange_center": workbook.add_format({
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": "#000000",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#F4B084",
            "num_format": "#,##0"
        }),
        "orange_left": workbook.add_format({
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": "#000000",
            "border": 1,
            "align": "left",
            "valign": "vcenter",
            "bg_color": "#F4B084",
            "num_format": "#,##0"
        }),
        "total_center": workbook.add_format({
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": "#000000",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#D9D9D9",
            "num_format": "#,##0"
        }),
        "total_left": workbook.add_format({
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": "#000000",
            "border": 1,
            "align": "left",
            "valign": "vcenter",
            "bg_color": "#D9D9D9",
            "num_format": "#,##0"
        }),
        "percent_center": workbook.add_format({
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": "#000000",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#D9D9D9",
            "num_format": "0.0%"
        }),
        "percent_left": workbook.add_format({
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": "#000000",
            "border": 1,
            "align": "left",
            "valign": "vcenter",
            "bg_color": "#D9D9D9",
            "num_format": "0.0%"
        })
    }


def write_legend(ws, row, formats):
    ws.merge_range(row, 11, row, 15, "KAWASAN MAJORITI MELAYU", formats["legend_melayu"])
    ws.merge_range(row + 1, 11, row + 1, 15, "KAWASAN MAJORITI BUKAN MELAYU", formats["legend_bukan"])


def is_bukan_melayu_majority(row):
    melayu = float(row.get("MELAYU", 0) or 0)
    cina = float(row.get("CINA", 0) or 0)
    india = float(row.get("INDIA", 0) or 0)
    lain_lain = float(row.get("LAIN-LAIN", 0) or 0)

    return max(cina, india, lain_lain) > melayu


def write_demografik_table(ws, start_row, title, subtitle, rows, name_header, formats):
    columns = [
        "KOD", "NAMA", "JUMLAH PENGUNDI",
        "MELAYU", "CINA", "INDIA", "LAIN-LAIN",
        "18-25", "26-40", "41-60", "61>",
        "L", "P",
        "PEMILIH", "POLIS", "ASKAR"
    ]

    data_df, total_row, percent_row = split_total_and_percent(rows)

    r = start_row

    ws.write(r, 0, title, formats["title"])
    r += 1

    ws.write(r, 0, subtitle, formats["subtitle"])

    write_legend(ws, r + 1, formats)

    r += 4

    ws.merge_range(r, 0, r + 1, 0, "KOD", formats["header"])
    ws.merge_range(r, 1, r + 1, 1, name_header, formats["header"])
    ws.merge_range(r, 2, r + 1, 2, "JUMLAH PENGUNDI", formats["header"])

    ws.merge_range(r, 3, r, 6, "KAUM", formats["header"])
    ws.merge_range(r, 7, r, 10, "PERINGKAT UMUR", formats["header"])
    ws.merge_range(r, 11, r, 12, "JANTINA", formats["header"])
    ws.merge_range(r, 13, r, 15, "NO TEL", formats["header"])

    r += 1

    for c, header in enumerate(columns[3:], start=3):
        ws.write(r, c, header, formats["header"])

    r += 1

    max_name_len = len(str(name_header))

    for _, row in data_df.iterrows():
        is_bukan_melayu = is_bukan_melayu_majority(row)
        max_name_len = max(max_name_len, len(str(row.get("NAMA", ""))))

        for c, col in enumerate(columns):
            value = row.get(col, "")

            if c == 1:
                fmt = formats["orange_left"] if is_bukan_melayu else formats["left"]
            else:
                fmt = formats["orange_center"] if is_bukan_melayu else formats["center"]

            ws.write(r, c, value, fmt)

        r += 1

    r += 1

    if total_row:
        ws.write(r, 1, "JUMLAH PENGUNDI", formats["total_left"])

        for c, col in enumerate(columns[2:], start=2):
            ws.write(r, c, total_row.get(col, ""), formats["total_center"])

        r += 1

    if percent_row:
        ws.write(r, 1, "PERATUSAN", formats["percent_left"])
        ws.write(r, 2, "", formats["percent_center"])

        for c, col in enumerate(columns[3:], start=3):
            ws.write(r, c, percent_row.get(col, 0), formats["percent_center"])

        r += 1

    ws.set_h_pagebreaks([r])

    return r + 2, max_name_len


def write_demografik_xlsx_bytes(raw_df, cleaned_df, base_name):
    raw, tel_summary = prepare_demografik_df(raw_df, cleaned_df)

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        sheet = workbook.add_worksheet("DEMOGRAFIK")
        writer.sheets["DEMOGRAFIK"] = sheet

        formats = make_formats(workbook)

        sheet.set_column(0, 0, 8)
        sheet.set_column(1, 1, 20)
        sheet.set_column(2, 2, 20)
        sheet.set_column(3, 15, 9)

        parl_label = format_parlimen_label(raw, base_name)

        current_row = 0
        max_b_width = 20

        parlimen_rows = []

        for (kod_dun, nama_dun), grp in raw.groupby(["kod_dun", "nama_dun"], dropna=False):
            dun_code = format_dun_code(kod_dun)
            dun_name = str(nama_dun).strip().upper()

            tel_counts = get_tel_counts_for_group(
                tel_summary,
                {
                    "kod_dun": kod_dun,
                    "nama_dun": nama_dun
                }
            )

            parlimen_rows.append(
                build_demo_row(dun_code, dun_name, grp, tel_counts)
            )

        current_row, section_b_width = write_demografik_table(
            sheet,
            current_row,
            f"DEMOGRAFIK {parl_label}",
            parl_label,
            parlimen_rows,
            "DUN",
            formats
        )

        max_b_width = max(max_b_width, section_b_width + 2)

        for (kod_dun, nama_dun), dun_grp in raw.groupby(["kod_dun", "nama_dun"], dropna=False):
            dun_code = format_dun_code(kod_dun)
            dun_name = str(nama_dun).strip().upper()

            dun_label = f"{parl_label} - {dun_code} {dun_name}".strip()

            dm_rows = []

            for (kod_dm, nama_dm), dm_grp in dun_grp.groupby(["kod_dm", "nama_dm"], dropna=False):
                dm_code = format_dm_code(kod_dm)
                dm_name = str(nama_dm).strip().upper()

                tel_counts = get_tel_counts_for_group(
                    tel_summary,
                    {
                        "kod_dun": kod_dun,
                        "nama_dun": nama_dun,
                        "kod_dm": kod_dm,
                        "nama_dm": nama_dm
                    }
                )

                dm_rows.append(
                    build_demo_row(dm_code, dm_name, dm_grp, tel_counts)
                )

            current_row, section_b_width = write_demografik_table(
                sheet,
                current_row,
                f"DEMOGRAFIK {parl_label}",
                dun_label,
                dm_rows,
                "DM",
                formats
            )

            max_b_width = max(max_b_width, section_b_width + 2)

        sheet.set_column(1, 1, max(20, min(max_b_width, 60)))

    output.seek(0)

    return output.getvalue()
