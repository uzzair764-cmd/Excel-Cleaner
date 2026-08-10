import io
import re
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo

st.set_page_config(page_title="Demografik Generator", layout="wide")
st.title("DEMOGRAFIK Generator")

uploaded_files = st.file_uploader("Upload Excel file(s)", type=["xlsx", "xls"], accept_multiple_files=True)

HEADERS = [
    'KOD DM', 'NAMA DM', 'JUMLAH',
    'LELAKI', 'LELAKI (%)', 'PEREMPUAN', 'PEREMPUAN (%)',
    'MELAYU', 'MELAYU (%)', 'CINA', 'CINA (%)', 'INDIA', 'INDIA (%)', 'LAIN-LAIN', 'LAIN-LAIN (%)',
    '18-21', '18-21 (%)', '22-30', '22-30 (%)', '31-40', '31-40 (%)',
    '41-50', '41-50 (%)', '51-60', '51-60 (%)', '61+', '61+ (%)',
    'UMNO', 'UMNO (%)', 'PKR', 'PKR (%)', 'PAS', 'PAS (%)', 'PPBM', 'PPBM (%)',
    'PUTIH', 'PUTIH (%)', 'KELABU', 'KELABU (%)', 'HITAM', 'HITAM (%)',
    'PENGUNDI AWAL', 'PENGUNDI AWAL (%)', 'POLIS', 'POLIS (%)',
    'PASANGAN POLIS', 'PASANGAN POLIS (%)', 'ASKAR', 'ASKAR (%)',
    'PASANGAN ASKAR', 'PASANGAN ASKAR (%)'
]

MAIN_RACES = ['MELAYU', 'CINA', 'INDIA', 'LAIN-LAIN']
AGE_GROUPS = ['18-21', '22-30', '31-40', '41-50', '51-60', '61+']
PARTY_COLS = ['UMNO', 'PKR', 'PAS', 'PPBM']
SIKAP_COLS = ['PUTIH', 'KELABU', 'HITAM']


def get_col(df, possible_names):
    col_map = {str(c).lower().strip(): c for c in df.columns}
    for name in possible_names:
        if name.lower().strip() in col_map:
            return col_map[name.lower().strip()]
    return None


def clean_service_no(value):
    n = str(value).strip().upper()
    return '' if n in {'', 'NAN', 'NONE', 'NULL'} else n


def clean_filename(value):
    name = str(value).strip().upper()
    name = re.sub(r'[\\/:*?"<>|]', ' ', name)
    name = ' '.join(name.split())
    return name if name else 'OUTPUT'


def kod_dun_digits(value):
    return re.sub(r'\D', '', str(value).strip().split('.')[0])


def kod_dun_sort_key(value):
    digits = kod_dun_digits(value)
    return int(digits) if digits else -1


def format_kod_dm(value):
    kod = str(value).strip()
    if kod in {'', 'None', 'nan', 'NaN'}:
        return ''
    kod = kod.split('.')[0].zfill(7)
    return f"{kod[:3]}/{kod[3:5]}/{kod[5:]}"


def normalise_race(value):
    r = str(value).strip().upper()
    return r if r in {'MELAYU', 'CINA', 'INDIA'} else 'LAIN-LAIN'


def normalise_gender(value):
    g = str(value).strip().upper()
    if g in {'L', 'LELAKI', 'MALE', 'M'}:
        return 'L'
    if g in {'P', 'PEREMPUAN', 'FEMALE', 'F'}:
        return 'P'
    return ''


def normalise_sikap(value):
    s = str(value).strip().upper()
    if s in {'KELABU-LAMA', 'KELABU-BARU'}:
        return 'KELABU'
    return s if s in {'PUTIH', 'KELABU', 'HITAM'} else ''


def classify_awal(value):
    n = clean_service_no(value)
    if not n:
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
    return clean_service_no(value).startswith('T')


def get_age_group(value):
    try:
        age = int(float(value))
        if 18 <= age <= 21: return '18-21'
        if 22 <= age <= 30: return '22-30'
        if 31 <= age <= 40: return '31-40'
        if 41 <= age <= 50: return '41-50'
        if 51 <= age <= 60: return '51-60'
        if age >= 61: return '61+'
    except Exception:
        pass
    return ''


def pct(part, total):
    return round(part / total * 100, 1) if total else 0


def build_dm_row(kod_dm, nama_dm, grp):
    total = len(grp)
    row = {'KOD DM': format_kod_dm(kod_dm), 'NAMA DM': nama_dm, 'JUMLAH': total}
    sex_vc = grp['_jantina'].value_counts(); race_vc = grp['_race'].value_counts(); age_vc = grp['_age_group'].value_counts()
    party_vc = grp['_party'].value_counts(); sikap_vc = grp['_sikap'].value_counts(); awal_vc = grp['_awal_type'].value_counts()
    for key, label in [('L', 'LELAKI'), ('P', 'PEREMPUAN')]:
        c=int(sex_vc.get(key,0)); row[label]=c; row[f'{label} (%)']=pct(c,total)
    for race in MAIN_RACES:
        c=int(race_vc.get(race,0)); row[race]=c; row[f'{race} (%)']=pct(c,total)
    for age in AGE_GROUPS:
        c=int(age_vc.get(age,0)); row[age]=c; row[f'{age} (%)']=pct(c,total)
    for party in PARTY_COLS:
        c=int(party_vc.get(party,0)); row[party]=c; row[f'{party} (%)']=pct(c,total)
    for sikap in SIKAP_COLS:
        c=int(sikap_vc.get(sikap,0)); row[sikap]=c; row[f'{sikap} (%)']=pct(c,total)
    pengundi_awal=int(grp['_NoPerkhidmatan_clean'].ne('').sum()); polis=int(awal_vc.get('POLIS',0)); askar=int(awal_vc.get('ASKAR',0))
    pasangan_polis=int(grp['_Pasangan Polis'].sum()); pasangan_askar=int(grp['_Pasangan Askar'].sum())
    row['PENGUNDI AWAL']=pengundi_awal; row['PENGUNDI AWAL (%)']=pct(pengundi_awal,total)
    row['POLIS']=polis; row['POLIS (%)']=pct(polis,total); row['PASANGAN POLIS']=pasangan_polis; row['PASANGAN POLIS (%)']=pct(pasangan_polis,total)
    row['ASKAR']=askar; row['ASKAR (%)']=pct(askar,total); row['PASANGAN ASKAR']=pasangan_askar; row['PASANGAN ASKAR (%)']=pct(pasangan_askar,total)
    return row


def add_total_row(df):
    total=int(df['JUMLAH'].sum()); row={'KOD DM':'','NAMA DM':'','JUMLAH':total}
    for h in HEADERS:
        if h in {'KOD DM','NAMA DM','JUMLAH'}: continue
        if h.endswith('(%)'):
            base=h.replace(' (%)',''); row[h]=pct(row.get(base,0),total)
        else:
            row[h]=pd.to_numeric(df[h],errors='coerce').fillna(0).sum()
    return pd.concat([df,pd.DataFrame([row])],ignore_index=True)


def build_rumusan_df(dun_df):
    rows=[]
    for (kod_dm,nama_dm),grp in dun_df.groupby(['_KOD DM','_NAMA DM'],dropna=False): rows.append(build_dm_row(kod_dm,nama_dm,grp))
    result=pd.DataFrame(rows)
    for h in HEADERS:
        if h not in result.columns: result[h]=0
    return add_total_row(result[HEADERS].sort_values(by='KOD DM',kind='stable'))


def add_parliament_total_row(final_df,kod_parlimen,nama_parlimen):
    row=build_dm_row(kod_parlimen,nama_parlimen,final_df); row['KOD DM']=str(kod_parlimen).strip(); row['NAMA DM']=str(nama_parlimen).strip().upper()
    return pd.DataFrame([row])[HEADERS]


def build_age_race_gender_rows(dun_df):
    rows=[]
    for race in MAIN_RACES:
        race_df=dun_df[dun_df['_race']==race]
        header={'LABEL':race}
        for age in AGE_GROUPS:
            header[age]=age; header[f'{age} (%)']=f'{age} (%)'
        rows.append(('header',header))
        for gender_code,label in [('L','Lelaki'),('P','Perempuan')]:
            gender_df=race_df[race_df['_jantina']==gender_code]; total=len(gender_df); data={'LABEL':label}
            for age in AGE_GROUPS:
                count=int((gender_df['_age_group']==age).sum()); data[age]=count; data[f'{age} (%)']=pct(count,total)
            rows.append(('gender',data))
    return rows


def main_group_style():
    BLUE='9DC3E6'; GREEN='A9D18E'; ORANGE='F4B183'; YELLOW='FFD966'; PURPLE='B4A7D6'; WHITE='D9D9D9'; fills={}
    for c in range(4,8): fills[c]=BLUE
    for c in range(8,16): fills[c]=GREEN
    for c in range(16,28): fills[c]=ORANGE
    for c in range(28,36): fills[c]=YELLOW
    for c in range(36,42): fills[c]=WHITE
    for c in range(42,52): fills[c]=PURPLE
    return fills,{1,4,8,16,28,36,42,44,46,48,50},{2,7,15,27,35,41,43,45,47,49,51}


def write_main_table(ws,combined_df,total_rows,parliament_row):
    fills,left_edges,right_edges=main_group_style(); thin=Side(style='thin',color='000000'); medium=Side(style='medium',color='000000')
    for c,h in enumerate(HEADERS,1):
        cell=ws.cell(1,c,h); cell.font=Font(name='Calibri',size=11,bold=True); cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); cell.fill=PatternFill('solid',fgColor=fills[c]) if c in fills else PatternFill(fill_type=None)
        cell.border=Border(left=medium if c in left_edges else thin,right=medium if c in right_edges else thin,top=medium,bottom=medium)
    for r,values in enumerate(combined_df.itertuples(index=False),2):
        idx=r-2; special=idx in total_rows or idx==parliament_row
        for c,value in enumerate(values,1):
            cell=ws.cell(r,c,value); cell.font=Font(name='Calibri',size=11,bold=special); cell.alignment=Alignment(horizontal='left' if c==2 else 'center',vertical='center')
            if c in fills and special: cell.fill=PatternFill('solid',fgColor=fills[c])
            cell.border=Border(left=medium if c in left_edges else thin,right=medium if c in right_edges else thin,top=medium if special else thin,bottom=medium if special else thin)
            if isinstance(value,(int,float)): cell.number_format='0.0' if '(%)' in HEADERS[c-1] else '#,##0'
    widths={'A':13,'B':25,'C':13,'D':11.3,'E':14.7,'F':17,'G':20.6,'H':13.1,'I':16.6,'J':10,'K':13.4,'L':10.7,'M':14.1,'N':14.6,'O':18.1,'P':10.3,'Q':13.7,'R':10.3,'S':13.7,'T':10.3,'U':13.7,'V':10.3,'W':13.7,'X':10.3,'Y':13.7,'Z':8.6,'AA':12,'AB':9,'AC':12.4,'AD':9,'AE':12.4,'AF':10.9,'AG':14.3,'AH':11.7,'AI':15.1,'AJ':11,'AK':14.4,'AL':12.4,'AM':15.9,'AN':11.6,'AO':15,'AP':21.3,'AQ':24.9,'AR':10.6,'AS':14,'AT':21.4,'AU':25,'AV':11.4,'AW':14.9,'AX':22.4,'AY':26}
    for col,width in widths.items(): ws.column_dimensions[col].width=width
    tab=Table(displayName='Table_Demografik_Combined',ref=f'A1:AY{len(combined_df)+1}'); tab.tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showFirstColumn=False,showLastColumn=False,showRowStripes=False,showColumnStripes=False); ws.add_table(tab)


def write_dun_age_table(ws,start_row,kod_dun,nama_dun,dun_df):
    thin=Side(style='thin',color='000000'); medium=Side(style='medium',color='000000'); green='A9D18E'; orange='F4B183'; blue='9DC3E6'; end_col=13
    title=ws.cell(start_row,1,f"N.{kod_dun_digits(kod_dun)[-2:].zfill(2)} {str(nama_dun).upper()} - UMUR MENGIKUT KAUM DAN JANTINA")
    ws.merge_cells(start_row=start_row,start_column=1,end_row=start_row,end_column=end_col); title.font=Font(name='Calibri',size=14,bold=True); title.alignment=Alignment(horizontal='center',vertical='center'); title.fill=PatternFill('solid',fgColor='D9EAD3'); title.border=Border(left=medium,right=medium,top=medium,bottom=medium); ws.row_dimensions[start_row].height=24
    current=start_row+1
    for row_type,data in build_age_race_gender_rows(dun_df):
        values=[data['LABEL']]
        for age in AGE_GROUPS: values += [data[age],data[f'{age} (%)']]
        for c,value in enumerate(values,1):
            cell=ws.cell(current,c,value); cell.font=Font(name='Calibri',size=11,bold=row_type=='header'); cell.alignment=Alignment(horizontal='left' if c==1 else 'center',vertical='center',wrap_text=True)
            if row_type=='header': cell.fill=PatternFill('solid',fgColor=green if c==1 else orange); top=bottom=medium
            else: cell.fill=PatternFill('solid',fgColor='FFFFFF' if c==1 else blue); top=bottom=thin
            cell.border=Border(left=medium if c==1 else thin,right=medium if c==end_col else thin,top=top,bottom=bottom)
            if row_type=='gender' and c>1 and isinstance(value,(int,float)): cell.number_format='0.0' if c in {3,5,7,9,11,13} else '#,##0'
        current += 1
    widths={'A':18,'B':11,'C':13.7,'D':11,'E':13.7,'F':11,'G':13.7,'H':11,'I':13.7,'J':11,'K':13.7,'L':8.6,'M':13.7}
    for col,width in widths.items(): ws.column_dimensions[col].width=max(ws.column_dimensions[col].width or 0,width)
    return current+1


def generate_demografik(uploaded_files):
    all_data=[]; logs=[]
    for uploaded_file in uploaded_files:
        fname=uploaded_file.name
        try:
            df=pd.read_excel(uploaded_file,dtype=str); df.columns=[str(c).strip() for c in df.columns]
            col_dm=get_col(df,['KOD DM','kod_dm']); col_nama_dm=get_col(df,['NamaDM','nama_dm','NAMA DM']); col_kod_dun=get_col(df,['kod_dun','KOD DUN','KODDUN']); col_nama_dun=get_col(df,['nama_dun','DUN','NAMA DUN']); col_kod_parlimen=get_col(df,['kod_parlimen','KOD PARLIMEN','KODPARLIMEN']); col_nama_parlimen=get_col(df,['nama_parlimen','NAMA PARLIMEN','NAMAPARLIMEN']); col_jantina=get_col(df,['JANTINA','jantina']); col_bangsa=get_col(df,['kaum','BANGSA','kategori_kaum']); col_umur=get_col(df,['UMUR','umur']); col_party=get_col(df,['party','PARTY']); col_sikap=get_col(df,['CATATAN','sikap']); col_no=get_col(df,['NoPerkhidmatan','noperkhidmatan']); col_pasangan=get_col(df,['NoKPPasangan','NoPerkhidmatanPasangan','noperkhidmatanpasangan'])
            required={'KOD DM':col_dm,'NamaDM':col_nama_dm,'kod_dun':col_kod_dun,'nama_dun / DUN':col_nama_dun,'kod_parlimen':col_kod_parlimen,'nama_parlimen':col_nama_parlimen,'JANTINA':col_jantina,'BANGSA':col_bangsa,'UMUR':col_umur,'NoPerkhidmatan':col_no,'NoKPPasangan':col_pasangan}; missing=[k for k,v in required.items() if v is None]
            if missing: logs.append(f"Skipped {fname} — missing columns: {missing}"); continue
            df['_KOD_DUN']=df[col_kod_dun].fillna('').astype(str).str.strip(); df['_NAMA_DUN']=df[col_nama_dun].fillna('').astype(str).str.strip().str.upper(); df['_KOD_PARLIMEN']=df[col_kod_parlimen].fillna('').astype(str).str.strip(); df['_NAMA_PARLIMEN']=df[col_nama_parlimen].fillna('').astype(str).str.strip().str.upper(); df['_KOD DM']=df[col_dm].fillna('').astype(str).str.strip(); df['_NAMA DM']=df[col_nama_dm].fillna('').astype(str).str.strip(); df['_jantina']=df[col_jantina].fillna('').apply(normalise_gender); df['_race']=df[col_bangsa].apply(normalise_race); df['_age_group']=df[col_umur].apply(get_age_group); df['_party']=df[col_party].fillna('').astype(str).str.strip().str.upper() if col_party else ''; df['_sikap']=df[col_sikap].apply(normalise_sikap) if col_sikap else ''; df['_NoPerkhidmatan_clean']=df[col_no].apply(clean_service_no); df['_NoKPPasangan_clean']=df[col_pasangan].apply(clean_service_no); df['_awal_type']=df['_NoPerkhidmatan_clean'].apply(classify_awal); df['_Pasangan Polis']=df['_NoKPPasangan_clean'].apply(lambda x:1 if is_polis(x) else 0); df['_Pasangan Askar']=df['_NoKPPasangan_clean'].apply(lambda x:1 if is_askar(x) else 0)
            all_data.append(df); logs.append(f"Loaded {fname}: {len(df):,} rows")
        except Exception as e: logs.append(f"Error reading {fname}: {e}")
    if not all_data: raise ValueError("No valid data loaded.\n"+"\n".join(logs))
    final_df=pd.concat(all_data,ignore_index=True)
    combos=final_df[['_KOD_PARLIMEN','_NAMA_PARLIMEN']].drop_duplicates(); combos=combos[(combos['_KOD_PARLIMEN']!='')&(combos['_NAMA_PARLIMEN']!='')]
    if len(combos)!=1: raise ValueError(f"Expected exactly one Parliament in the uploaded data, found {len(combos)}.")
    kod_parlimen,nama_parlimen=combos.iloc[0].tolist()
    dun_combos=final_df[['_KOD_DUN','_NAMA_DUN']].drop_duplicates(); dun_combos=dun_combos[(dun_combos['_KOD_DUN']!='')&(dun_combos['_NAMA_DUN']!='')]; dun_combos=sorted(dun_combos.itertuples(index=False,name=None),key=lambda x:kod_dun_sort_key(x[0]))
    if not dun_combos: raise ValueError('No DUN name/kod_dun found in the uploaded data.')
    combined_parts=[]; total_rows=set(); dun_data=[]
    for kod_dun,nama_dun in dun_combos:
        dun_df=final_df[(final_df['_KOD_DUN']==kod_dun)&(final_df['_NAMA_DUN']==nama_dun)]; rumusan=build_rumusan_df(dun_df); start=sum(len(x) for x in combined_parts); total_rows.add(start+len(rumusan)-1); combined_parts.append(rumusan); dun_data.append((kod_dun,nama_dun,dun_df.copy())); logs.append(f"Built DUN N.{kod_dun_digits(kod_dun)[-2:].zfill(2)} {nama_dun}: {len(dun_df):,} rows, {len(rumusan)-1} DM(s)")
    parliament_df=add_parliament_total_row(final_df,kod_parlimen,nama_parlimen); combined_df=pd.concat(combined_parts+[parliament_df],ignore_index=True); parliament_row=len(combined_df)-1
    wb=Workbook(); ws=wb.active; ws.title='DEMOGRAFIK'; write_main_table(ws,combined_df,total_rows,parliament_row)
    next_row=len(combined_df)+3
    for kod_dun,nama_dun,dun_df in dun_data:
        next_row=write_dun_age_table(ws,next_row,kod_dun,nama_dun,dun_df); logs.append(f"Added age group per kaum and jantina for N.{kod_dun_digits(kod_dun)[-2:].zfill(2)} {nama_dun}")
    ws.freeze_panes='A2'; output=io.BytesIO(); wb.save(output); output.seek(0); out_name=f"DEMOGRAFIK {clean_filename(nama_parlimen)}.xlsx"
    logs.append(f"Combined {len(dun_combos)} DUN(s) into one worksheet"); logs.append(f"Parliament grand total: {kod_parlimen} - {nama_parlimen} ({len(final_df):,} rows)"); return output.getvalue(),out_name,logs


if uploaded_files:
    if st.button('Generate DEMOGRAFIK'):
        try:
            excel_bytes,out_name,logs=generate_demografik(uploaded_files); st.success(f'Generated: {out_name}')
            with st.expander('Processing log'):
                for log in logs: st.write(log)
            st.download_button('Download Excel',data=excel_bytes,file_name=out_name,mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        except Exception as e: st.error(str(e))
else:
    st.info('Upload one or more Excel files to start.')
