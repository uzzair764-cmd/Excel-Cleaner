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


def prepare_demografik_df(df):
    df = df.copy()

    needed_cols = [
        "nokp", "umur", "jantina", "kaum_spr", "kategori_kaum",
        "party", "number", "kod_parlimen", "nama_parlimen",
        "kod_dun", "nama_dun", "kod_dm", "nama_dm"
    ]

    for col in needed_cols:
        if col not in df.columns:
            df[col] = ""

    race_source = "kategori_kaum" if "kategori_kaum" in df.columns else "kaum_spr"

    df["_race"] = df[race_source].apply(normalise_race)
    df["_age"] = df["umur"].apply(age_group)
    df["_jantina"] = df["jantina"].astype(str).str.strip().str.upper()

    raw_number = df["number"].astype(str).str.strip()
    df["_pemilih_tel"] = raw_number.ne("") & raw_number.str.lower().ne("nan")

    service_col = get_service_col(df)

    if service_col:
        svc = df[service_col].astype(str).str.strip().str.upper()
    else:
        svc = pd.Series([""] * len(df), index=df.index)

    df["_polis"] = svc.str.startswith(("G", "R"), na=False)
    df["_askar"] = svc.str.startswith("T", na=False)

    return df


def build_demo_row(code, name, group):
    total = len(group)

    race_vc = group["_race"].value_counts()
    age_vc = group["_age"].value_counts()
    gender_vc = group["_jantina"].value_counts()

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

    row["PEMILIH"] = int(group["_pemilih_tel"].sum())
    row["POLIS"] = int(group["_polis"].sum())
    row["ASKAR"] = int(group["_askar"].sum())

    return row


def add_total_and_percent_rows(rows):
    df = pd.DataFrame(rows)

    if df.empty:
        return df

    count_cols = ["JUMLAH PENGUNDI"] + RACES + AGE_GROUPS + GENDERS + TEL_COLS

    total_row = {
        "KOD": "",
        "NAMA": "JUMLAH PENGUNDI"
    }

    percent_row = {
        "KOD": "",
        "NAMA": "PERATUSAN"
    }

    for col in count_cols:
        total_row[col] = int(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())

    grand_total = total_row["JUMLAH PENGUNDI"]

    percent_row["JUMLAH PENGUNDI"] = ""

    for col in RACES + AGE_GROUPS + GENDERS + TEL_COLS:
        percent_row[col] = pct(total_row[col], grand_total)

    return pd.concat([df, pd.DataFrame([total_row, percent_row])], ignore_index=True)


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


def write_legend(ws, workbook):
    legend_melayu_fmt = workbook.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "bold": True,
        "font_color": "#000000",
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "bg_color": "#FFFFFF"
    })

    legend_bukan_fmt = workbook.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "bold": True,
        "font_color": "#000000",
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "bg_color": "#F4B084"
    })

    ws.merge_range("L3:P3", "KAWASAN MAJORITI MELAYU", legend_melayu_fmt)
    ws.merge_range("L4:P4", "KAWASAN MAJORITI BUKAN MELAYU", legend_bukan_fmt)


def write_demografik_table(ws, workbook, start_row, title, subtitle, table_df, name_header):
    title_fmt = workbook.add_format({
        "font_name": "Calibri",
        "font_size": 22,
        "bold": True,
        "font_color": "#000000",
        "align": "left",
        "valign": "vcenter"
    })

    subtitle_fmt = workbook.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "bold": True,
        "font_color": "#000000",
        "align": "left",
        "valign": "vcenter"
    })

    header_fmt = workbook.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "bold": True,
        "font_color": "#000000",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "bg_color": "#D9D9D9"
    })

    cell_center_fmt = workbook.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "bold": True,
        "font_color": "#000000",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0"
    })

    cell_left_fmt = workbook.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "bold": True,
        "font_color": "#000000",
        "border": 1,
        "align": "left",
        "valign": "vcenter",
        "num_format": "#,##0"
    })

    orange_center_fmt = workbook.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "bold": True,
        "font_color": "#000000",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "bg_color": "#F4B084",
        "num_format": "#,##0"
    })

    orange_left_fmt = workbook.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "bold": True,
        "font_color": "#000000",
        "border": 1,
        "align": "left",
        "valign": "vcenter",
        "bg_color": "#F4B084",
        "num_format": "#,##0"
    })

    total_center_fmt = workbook.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "bold": True,
        "font_color": "#000000",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "bg_color": "#D9D9D9",
        "num_format": "#,##0"
    })

    total_left_fmt = workbook.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "bold": True,
        "font_color": "#000000",
        "border": 1,
        "align": "left",
        "valign": "vcenter",
        "bg_color": "#D9D9D9",
        "num_format": "#,##0"
    })

    percent_center_fmt = workbook.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "bold": True,
        "font_color": "#000000",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "bg_color": "#D9D9D9",
        "num_format": "0.0%"
    })

    percent_left_fmt = workbook.add_format({
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

    columns = [
        "KOD", "NAMA", "JUMLAH PENGUNDI",
        "MELAYU", "CINA", "INDIA", "LAIN-LAIN",
        "18-25", "26-40", "41-60", "61>",
        "L", "P",
        "PEMILIH", "POLIS", "ASKAR"
    ]

    r = start_row

    ws.merge_range(r, 0, r, 2, title, title_fmt)
    r += 1

    ws.merge_range(r, 0, r, 1, subtitle, subtitle_fmt)
    r += 4

    ws.merge_range(r, 0, r + 1, 0, "KOD", header_fmt)
    ws.merge_range(r, 1, r + 1, 1, name_header, header_fmt)
    ws.merge_range(r, 2, r + 1, 2, "JUMLAH PENGUNDI", header_fmt)

    ws.merge_range(r, 3, r, 6, "KAUM", header_fmt)
    ws.merge_range(r, 7, r, 10, "PERINGKAT UMUR", header_fmt)
    ws.merge_range(r, 11, r, 12, "JANTINA", header_fmt)
    ws.merge_range(r, 13, r, 15, "NO TEL", header_fmt)

    r += 1

    for c, header in enumerate(columns[3:], start=3):
        ws.write(r, c, header, header_fmt)

    r += 1

    for _, row in table_df.iterrows():
        name_value = str(row.get("NAMA", "")).upper()

        is_total = name_value == "JUMLAH PENGUNDI"
        is_percent = name_value == "PERATUSAN"

        melayu = float(row.get("MELAYU", 0) or 0)
        non_melayu = (
            float(row.get("CINA", 0) or 0) +
            float(row.get("INDIA", 0) or 0) +
            float(row.get("LAIN-LAIN", 0) or 0)
        )

        is_bukan_melayu = (
            not is_total
            and not is_percent
            and non_melayu > melayu
        )

        for c, col in enumerate(columns):
            value = row.get(col, "")

            if c == 1:
                if is_percent:
                    fmt = percent_left_fmt
                elif is_total:
                    fmt = total_left_fmt
                elif is_bukan_melayu:
                    fmt = orange_left_fmt
                else:
                    fmt = cell_left_fmt
            else:
                if is_percent:
                    fmt = percent_center_fmt
                elif is_total:
                    fmt = total_center_fmt
                elif is_bukan_melayu:
                    fmt = orange_center_fmt
                else:
                    fmt = cell_center_fmt

            ws.write(r, c, value, fmt)

        r += 1

    return r + 2


def write_demografik_xlsx_bytes(raw_df, base_name):
    df = prepare_demografik_df(raw_df)

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        sheet = workbook.add_worksheet("DEMOGRAFIK")
        writer.sheets["DEMOGRAFIK"] = sheet

        sheet.set_column(0, 0, 12)
        sheet.set_column(1, 1, 36)
        sheet.set_column(2, 15, 14)

        parl_label = format_parlimen_label(df, base_name)

        write_legend(sheet, workbook)

        current_row = 0

        parlimen_rows = []

        for (kod_dun, nama_dun), grp in df.groupby(["kod_dun", "nama_dun"], dropna=False):
            dun_code = format_dun_code(kod_dun)
            dun_name = str(nama_dun).strip().upper()

            parlimen_rows.append(
                build_demo_row(dun_code, dun_name, grp)
            )

        parlimen_table = add_total_and_percent_rows(parlimen_rows)

        current_row = write_demografik_table(
            sheet,
            workbook,
            current_row,
            f"DEMOGRAFIK {parl_label}",
            parl_label,
            parlimen_table,
            "DUN"
        )

        for (kod_dun, nama_dun), dun_grp in df.groupby(["kod_dun", "nama_dun"], dropna=False):
            dun_code = format_dun_code(kod_dun)
            dun_name = str(nama_dun).strip().upper()

            dun_label = f"{parl_label} - {dun_code} {dun_name}".strip()

            dm_rows = []

            for (kod_dm, nama_dm), dm_grp in dun_grp.groupby(["kod_dm", "nama_dm"], dropna=False):
                dm_code = format_dm_code(kod_dm)
                dm_name = str(nama_dm).strip().upper()

                dm_rows.append(
                    build_demo_row(dm_code, dm_name, dm_grp)
                )

            dm_table = add_total_and_percent_rows(dm_rows)

            current_row = write_demografik_table(
                sheet,
                workbook,
                current_row,
                f"DEMOGRAFIK {parl_label}",
                dun_label,
                dm_table,
                "DM"
            )

    output.seek(0)

    return output.getvalue()
