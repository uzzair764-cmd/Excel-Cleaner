import io
import os
import re
import zipfile
import pandas as pd

from processors.demografik_processor import write_demografik_xlsx_bytes


FINAL_COLUMNS = [
    "id", "nokp", "name", "umur", "jantina", "kaum_spr", "number",
    "kod_lokaliti", "nama_lokaliti", "kod_dm", "nama_dm",
    "kod_dun", "nama_dun", "kod_parlimen", "nama_parlimen",
    "kod_negeri", "nama_negeri", "sikap", "party"
]

COLUMN_MAP = {
    "nama": "name",
    "phone 1": "number",
    "bangsa": "kaum_spr",
    "kaum": "kaum_spr",          # <-- added: maps new "kaum" header to kaum_spr
    "kod lokaliti": "kod_lokaliti",
    "lokaliti": "nama_lokaliti",
    "kod dm": "kod_dm",
    "dm": "nama_dm",
    "kod dun": "kod_dun",
    "dun": "nama_dun",
    "kod parlimen": "kod_parlimen",
    "parlimen": "nama_parlimen",
    "kelas": "sikap",
}


def parse_start_id(start_id):
    match = re.match(r"^(.*?)(\d+)$", str(start_id).strip())
    if not match:
        raise ValueError("ID must end with number. Example: CJ1000")
    return match.group(1), int(match.group(2)) + 1


def read_file(uploaded_file):
    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, dtype=str)

    if name.endswith(".xlsx"):
        return pd.read_excel(uploaded_file, dtype=str)

    raise ValueError("Only CSV and XLSX files are supported.")


def standardize_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    rename = {}

    for col in df.columns:
        key = col.lower().strip()

        if key in COLUMN_MAP:
            rename[col] = COLUMN_MAP[key]

        elif key == "negeri":
            rename[col] = "nama_negeri"

    return df.rename(columns=rename)


def clean_text(df):
    df = df.copy()

    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()

    return df


def clean_numbers(df):
    df = df.copy()
    total_rows = len(df)

    if "number" not in df.columns:
        df["number"] = ""

    df["number"] = (
        df["number"]
        .fillna("")
        .astype(str)
        .str.replace(r"\D", "", regex=True)
        .str.replace(r"^60", "0", regex=True)
    )

    invalid_mask = (
        (df["number"] == "") |
        (df["number"].str.contains("0000", na=False)) |
        (~df["number"].str.startswith("01", na=False)) |
        (~df["number"].str.len().isin([10, 11]))
    )

    invalid_removed = int(invalid_mask.sum())

    df = df.loc[~invalid_mask].copy()

    before_dedupe = len(df)
    df = df.drop_duplicates(subset=["number"], keep="first").copy()
    duplicate_removed = before_dedupe - len(df)

    return df.reset_index(drop=True), {
        "uploaded_rows": total_rows,
        "invalid_removed": invalid_removed,
        "duplicate_removed": duplicate_removed,
        "final_rows": len(df),
    }


def build_output_df(df, id_prefix, start_num):
    df = df.copy()

    df = df.drop(
        columns=["id", "umur2", "kategori_kaum"],
        errors="ignore"
    )

    row_count = len(df)

    df.insert(
        0,
        "id",
        [f"{id_prefix}{i}" for i in range(start_num, start_num + row_count)]
    )

    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    last_generated_id = f"{id_prefix}{start_num + row_count - 1}" if row_count else None

    return df[FINAL_COLUMNS], start_num + row_count, last_generated_id


def write_xlsx_bytes(df):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(
            writer,
            index=False,
            header=False,
            startrow=1,
            sheet_name="Sheet1"
        )

        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        header_fmt = workbook.add_format({
            "bold": False,
            "border": 0,
            "align": "center",
            "valign": "vcenter"
        })

        cell_fmt = workbook.add_format({
            "border": 0,
            "align": "center",
            "valign": "vcenter"
        })

        for col_idx, col_name in enumerate(df.columns):
            worksheet.write(0, col_idx, col_name, header_fmt)

        worksheet.set_column(
            0,
            len(df.columns) - 1,
            18,
            cell_fmt
        )

    output.seek(0)
    return output.getvalue()


def write_csv_chunks(df, chunk_size=50000):
    csv_df = pd.DataFrame({
        "id": df["id"].astype(str),
        "name": df["name"].astype(str),
        "number": "6" + df["number"].astype(str),
    })

    chunks = []

    for chunk_no, start in enumerate(range(0, len(csv_df), chunk_size), start=1):
        part = csv_df.iloc[start:start + chunk_size]

        csv_bytes = part.to_csv(
            index=False,
            encoding="utf-8-sig",
            lineterminator="\n"
        ).encode("utf-8-sig")

        chunks.append((chunk_no, csv_bytes))

    return chunks


def run_cleaner(
    uploaded_files,
    start_id,
    chunk_size=50000,
    remove_invalid=True,
    remove_duplicates=True,
    prefix_6=True
):
    id_prefix, current_num = parse_start_id(start_id)

    zip_buffer = io.BytesIO()
    summary = []

    overall_last_generated_id = None

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in uploaded_files:
            base_name = os.path.splitext(file.name)[0]

            raw_df = read_file(file)

            df = standardize_columns(raw_df)
            df = clean_text(df)

            cleaned_df, stats = clean_numbers(df)

            # DEMOGRAFIK logic:
            # - demographic columns use pre-cleaning df
            # - NO TEL columns use post-cleaning cleaned_df
            demografik_bytes = write_demografik_xlsx_bytes(
                raw_df=df,
                cleaned_df=cleaned_df,
                base_name=base_name
            )

            output_df, current_num, last_generated_id = build_output_df(
                cleaned_df,
                id_prefix,
                current_num
            )

            if last_generated_id:
                overall_last_generated_id = last_generated_id

            zipf.writestr(
                f"{base_name}/{base_name}.xlsx",
                write_xlsx_bytes(output_df)
            )

            zipf.writestr(
                f"{base_name}/DEMOGRAFIK {base_name}.xlsx",
                demografik_bytes
            )

            csv_chunks = write_csv_chunks(
                output_df,
                chunk_size=chunk_size
            )

            for chunk_no, csv_bytes in csv_chunks:
                zipf.writestr(
                    f"{base_name}/CSV/{base_name} {chunk_no}.csv",
                    csv_bytes
                )

            stats["file"] = file.name
            stats["csv_chunks"] = len(csv_chunks)
            stats["last_id_generated"] = last_generated_id or ""
            summary.append(stats)

    zip_buffer.seek(0)

    summary_df = pd.DataFrame(summary)

    if not summary_df.empty:
        summary_df.attrs["last_generated_id"] = overall_last_generated_id or ""

    return zip_buffer.getvalue(), summary_df
