import os
import pandas as pd
import sys

def main():
    output_filename = 'HASIL_MERGER_SEMUA.xlsx'
    
    current_dir = os.getcwd()
    
    files = [f for f in os.listdir(current_dir) if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$') and f != output_filename]

    if not files:
        print("--> Tidak ditemukan file Excel di folder ini")
        input("--> Tekan enter untuk keluar")
        return

    print(f"--> Ditemukan {len(files)} file untuk digabungkan")
    
    all_dataframes = []

    for filename in files:
        try:
            print(f"--> Membaca file: {filename}")
            df = pd.read_excel(filename)
            all_dataframes.append(df)
        except Exception as e:
            print(f"--> Gagal membaca {filename}: {e}")

    if all_dataframes:
        try:
            print("--> Sedang menggabungkan data...")
            merged_df = pd.concat(all_dataframes, ignore_index=True)
            
            print(f"--> Menyimpan hasil ke: {output_filename}")
            merged_df.to_excel(output_filename, index=False)
            
            print("--> Selesai. Semua data berhasil digabung")
            
        except Exception as e:
            print(f"--> Terjadi kesalahan saat menyimpan: {e}")
    else:
        print("--> Tidak ada data yang bisa digabungkan")

    input("--> Tekan enter untuk keluar")

if __name__ == "__main__":
    main()