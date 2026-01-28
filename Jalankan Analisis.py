import os
import shutil
import subprocess
import sys

def main():
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        input_dir = os.path.join(base_dir, 'Input')
        dapur_dir = os.path.join(base_dir, 'Dapur')

        print("--> Memulai proses...")

        if not os.path.exists(input_dir) or not os.path.exists(dapur_dir):
            print("--> Folder Input atau Dapur tidak ditemukan")
            input("--> Tekan enter untuk keluar")
            return

        files = [f for f in os.listdir(input_dir) if f.lower().endswith('.xls')]
        
        for f in files:
            shutil.copy2(os.path.join(input_dir, f), os.path.join(dapur_dir, f))
        
        print(f"--> Berhasil menyalin {len(files)} file ke Dapur")

        os.chdir(dapur_dir)

        print("--> Menjalankan convert_fakecxls.py...")
        subprocess.check_call([sys.executable, 'convert_fakecxls.py'])

        print("--> Menjalankan the_magic.py...")
        subprocess.check_call([sys.executable, 'the_magic.py'])

        output_filename = 'HASIL_ANALISIS_GABUNGAN.xlsx'
        
        if os.path.exists(output_filename):
            final_name = 'Hasil Analisis Pengiriman Pesanan vs Faktur Penjualan.xlsx'
            shutil.copy2(output_filename, os.path.join(base_dir, final_name))
            print(f"--> Sukses! File tersimpan sebagai: {final_name}")

            print("--> Membersihkan file xls dan xlsx di folder Dapur...")
            for f in os.listdir(dapur_dir):
                if f.lower().endswith(('.xls', '.xlsx')):
                    try:
                        os.remove(os.path.join(dapur_dir, f))
                    except OSError:
                        pass
            print("--> Pembersihan selesai")

        else:
            print("--> Gagal: File output tidak ditemukan")

    except subprocess.CalledProcessError as e:
        print(f"--> Gagal saat menjalankan script pendukung: {e}")
    except Exception as e:
        print(f"--> Terjadi kesalahan: {e}")

    input("--> Tekan enter untuk keluar")

if __name__ == '__main__':
    main()