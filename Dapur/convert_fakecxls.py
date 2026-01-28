import pandas as pd
import os
import glob
import warnings
import re
import html

warnings.filterwarnings("ignore")

def parse_accurate_xml_bruteforce(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()

        data_rows = []
        row_matches = re.findall(r'<Row.*?>(.*?)</Row>', content, re.DOTALL | re.IGNORECASE)

        for row_str in row_matches:
            cell_matches = re.findall(r'<Data[^>]*>(.*?)</Data>', row_str, re.DOTALL | re.IGNORECASE)
            cleaned_cells = [html.unescape(c.strip()) for c in cell_matches]
            if cleaned_cells:
                data_rows.append(cleaned_cells)

        if data_rows:
            return pd.DataFrame(data_rows)

    except Exception:
        return None
    return None

def clean_indo_number(value):
    if pd.isna(value) or value == "":
        return 0.0
    
    str_val = str(value)
    
    str_val = str_val.replace('(Dr)', '').replace('(Cr)', '').strip()
    
    if '(' in str_val and ')' in str_val:
        str_val = str_val.replace('(', '-').replace(')', '')
    
    str_val = str_val.replace('.', '')
    str_val = str_val.replace(',', '.')
    
    try:
        return float(str_val)
    except ValueError:
        return 0.0

def process_dataframe(df):
    if df is None or df.empty:
        return None

    start_row = 0
    for i in range(min(20, len(df))):
        non_empty = df.iloc[i].dropna().astype(str).str.strip().ne('').sum()
        if non_empty > 3: 
            start_row = i
            break
            
    df = df[start_row:]
    
    if len(df) > 0:
        df = df.iloc[1:]
    
    df.reset_index(drop=True, inplace=True)
    
    df = df.iloc[:, :7]
    
    while df.shape[1] < 7:
        df[df.shape[1]] = ""

    new_columns = ['Tanggal', 'Keterangan', 'Nomor Faktur', 'Detail', 'Debit', 'Kredit', 'Total']
    df.columns = new_columns

    cols_to_convert = ['Debit', 'Kredit', 'Total']
    
    for col in cols_to_convert:
        df[col] = df[col].apply(clean_indo_number)

    df.dropna(how='all', inplace=True)
    
    return df

def main():
    xls_files = glob.glob("*.xls")
    output_filename = "HASIL_EKSTRAK_GABUNGAN.xlsx"
    
    if not xls_files:
        print("--> Tidak ditemukan file .xls di folder ini.")
        return

    print(f"--> Menemukan {len(xls_files)} file. Memulai proses gabungan...")

    try:
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            processed_count = 0
            
            for file in xls_files:
                filename = os.path.basename(file)
                print(f"--> Memproses: {filename}")
                
                df = parse_accurate_xml_bruteforce(file)
                
                if df is not None:
                    df_clean = process_dataframe(df)
                    
                    if df_clean is not None and not df_clean.empty:
                        sheet_name = os.path.splitext(filename)[0][:31]
                        
                        invalid_chars = '[]:*?/\\'
                        for char in invalid_chars:
                            sheet_name = sheet_name.replace(char, '')
                            
                        df_clean.to_excel(writer, sheet_name=sheet_name, index=False)
                        print(f"--> Berhasil menambahkan sheet: {sheet_name}")
                        processed_count += 1
                    else:
                        print(f"--> Data kosong atau gagal dibersihkan: {filename}")
                else:
                    print(f"--> Gagal ekstrak konten: {filename}")
            
            if processed_count > 0:
                print(f"--> Selesai! File disimpan sebagai: {output_filename}")
            else:
                print("--> Tidak ada data yang berhasil dikonversi.")

    except Exception as e:
        print(f"--> Terjadi kesalahan saat menyimpan file: {e}")

if __name__ == "__main__":
    main()