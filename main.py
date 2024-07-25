import pandas as pd
import tkinter as tk
from tkinter import filedialog

def modify_excel_file(file_path):
    # Membaca file Excel
    df = pd.read_excel(file_path)

    # Mengonversi nama kolom menjadi string dan menghapus spasi berlebih
    df.columns = df.columns.str.strip()

    # Memeriksa apakah kolom yang diinginkan ada
    required_columns = ['C', 'D', 'E', 'F', 'G']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Kolom '{col}' tidak ditemukan dalam file Excel.")

    # Menambahkan teks sebagai prefix pada sel C2 hingga G46 dan menggabungkannya di sel H2 hingga H46
    for i in range(0, 46):  # baris 2 hingga 46 (0-based index untuk df)
        if i >= len(df):
            break
        c_value = '5 ' + str(df.at[i, 'C'])
        d_value = '4 ' + str(df.at[i, 'D'])
        e_value = '3 ' + str(df.at[i, 'E'])
        f_value = '2 ' + str(df.at[i, 'F'])
        g_value = '1 ' + str(df.at[i, 'G'])
        combined_value = f"{c_value}\n{d_value}\n{e_value}\n{f_value}\n{g_value}"
        df.at[i, 'H'] = combined_value

    # Menyimpan kembali file dengan prefix "fixed"
    output_file_path = 'fixed_' + file_path.split('/')[-1]
    df.to_excel(output_file_path, index=False)

    print(f"File telah disimpan sebagai {output_file_path}")

# Membuat jendela tkinter untuk dialog file
root = tk.Tk()
root.withdraw()  # Menyembunyikan jendela utama tkinter

file_path = filedialog.askopenfilename(
    title="Pilih file Excel",
    filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
)

if file_path:
    try:
        modify_excel_file(file_path)
    except ValueError as e:
        print(e)
else:
    print("Tidak ada file yang dipilih.")
