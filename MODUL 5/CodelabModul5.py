import pandas as pd
import os

file_name = "data_mahasiswa.xlsx"

if os.path.exists(file_name):
    try:
        df = pd.read_excel(file_name)
        data_mahasiswa = df.to_dict(orient="records")
    except Exception as e:
        print(f"Terjadi kesalahan membaca file Excel: {e}")
        data_mahasiswa = []
else:
    data_mahasiswa = []

print("Masukkan data mahasiswa. Ketik 'selesai' pada nama untuk mengakhiri.")

while True:
    nama = input("Masukkan Nama: ")
    if nama.lower() == "selesai":
        break

    if any(mahasiswa["Nama"] == nama for mahasiswa in data_mahasiswa):
        print("Nama sudah ada, masukkan nama yang berbeda.")
        continue

    semester = input("Masukkan Semester: ")
    mata_kuliah = input("Masukkan Mata Kuliah: ")

    data_mahasiswa.append({"Nama": nama, "Semester": semester, "Mata Kuliah": mata_kuliah})
    print("Data berhasil ditambahkan!")

if data_mahasiswa:
    df = pd.DataFrame(data_mahasiswa)
    try:
        df.to_excel(file_name, index=False)
        print(f"Data berhasil disimpan ke dalam file {file_name}")
    except Exception as e:
        print(f"Terjadi kesalahan saat menyimpan file Excel: {e}")
else:
    print("Tidak ada data yang disimpan.")
