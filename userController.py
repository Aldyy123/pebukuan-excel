from datetime import datetime
from colorama import Fore
from calendar import monthrange
import os
import pandas as pd

list_command = {
    1 : 'Membuat laporan harian',
    2: 'Menganalisa Data excel',
    3 : 'Grafik data bulanan'
}
def commandUserList():
    print('Selamat datang di pembukuan bisnis')
    print('Silahkan memilih command yang anda ingin lakukan')
    print('==================================')
    for i in list_command:
        print(str(i) +'. '+ list_command[i])
    print('==================================')
    return int(input('Silahkan masukan nomer perintah yang anda pilih: ') or 0)

        # Harian
        # Jika males ngisi tanggal default tanggal sekarang
        # Jika belum mengisi tanggal sebelumnya
        # Jika ingin meruntut dari tanggal ke tanggal berapa

def insert_file():
    # Menerima Nama file dari user yang nantinya akan di ganti ke extention exls
    # Mengecek suatu kondisi yang nantinya akan mengetahui bahwa file yang di tulis sudah ada atau belum
    # Jika belum maka akan menggunakan fungsi excel baru
    # Jika sudah ada maka akan menggunakan fungsi load workbok
    file = input('Masukan nama file yang sudah ada atau membuat yang baru: ')


    if len(file) <= 1:
       return insert_file()
    elif len(file) > 1:
        data = check_list_data_excel(file)
        print(data)
        return create_directory_check(file)
    else:
        print("error cuy")

def input_dates(msg = ''):
    tanggal = datetime.now().strftime("%d")
    year = datetime.now().strftime("%G")
    month  = datetime.now().strftime("%m")
    date_now  = datetime.now().strftime("%x")

    print("""
    \t Masukan tanggalan input untuk omset hariannya anda \n
    \t Tanggal bisa di atur sesuai keinginan anda \n
    \t Contoh 2-30 : {0} Untuk mengurutkan tanggalan sesuai input {1} \n
    \t Contoh 1,2,5,7,8 : {0} Untuk Memilih tanggal yang anda pilih saja {1}\n
    \t {0}Pilih tanggal yang spesifik jika anda ingin menambahkan satu tanggal{1} \n
    \t {3}{2}{1}
    """.format(Fore.GREEN, Fore.RESET, msg, Fore.RED))

    total_days = monthrange(int(year), int(month))[1]
    user_input_dates = input('Masukan tanggal, Default tanggal sekarang : ') or tanggal
    user_input_dates = check_tanggal(user_input_dates, total_days)
    return user_input_dates, date_now


def check_list_data_excel(file):
    try:
        file = f'./excel/{file}.xlsx'
        xl = pd.ExcelFile(file)
        print(xl.sheet_names)
        sheet_name = input('Masukan sheet yang ingin anda lihat: ')
        data = pd.read_excel(file, sheet_name=sheet_name)
        return data
    except Exception:
        pass


from helper import check_tanggal, create_directory_check