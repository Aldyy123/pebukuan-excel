from openpyxl import Workbook, styles, load_workbook, worksheet
from colorama import Fore
import re
from helper import store_data_omset

# Memasukan hasil excel untuk di olah datanya kemudian
# save data
# memasukan file excel yang sudah ada
# update data excel


def input_new_excel(name_file, dates, jam, omset, separator):
    wb = Workbook()
    ws = wb.active

    ws.merge_cells('A1:C2')
    ws['A1'] = 'Oktober'
    ws['A1'].alignment = styles.Alignment(
        horizontal='center', vertical='center')
    ws['A1'].font = styles.Font(size=30, bold=True)

    ws['A3'] = 'Tanggal'
    ws['B3'] = 'Omset'
    ws['C3'] = 'Waktu'

    store_data_omset(dates, ws, omset, jam, separator)
    
    wb.save('./excel/{}.xlsx'.format(name_file))


def update_excel(name_file, dates, jam, omset, separator):
    wb = load_workbook(f'./excel/{name_file}.xlsx')
    new_sheet = input("Apakah ingin membuat sheet baru? : ")
    ws = None
    if new_sheet in ['y']:
        new_sheet = input("Masukan nama sheet yang anda inginkan : ")
        ws = wb.create_sheet(new_sheet)
    else:
        chose_sheet = input("Masukan sheet yang ingin anda ubah : ")
        ws = wb[chose_sheet]

    store_data_omset(dates, ws, omset, jam, separator)

    wb.save('./excel/{}.xlsx'.format(name_file))


def input_omset_harian(tanggal_hari_ini, jam, alert=''):
    input_omset_today = input("""
    \t  Format jika anda memasukan lebih dari jumlah omset harian adalah
    \t  Example : {0} 2000, 90000, 10000 {1}
    \t  Sesuai jumlah tanggal yang anda input
    \t  {3}{2}{1}
    \t  Silahkan masukan omset hari ini: 
""".format(Fore.GREEN, Fore.RESET, alert, Fore.RED))
    omset_tunggal = input_omset_today
    cek_string = re.findall('[a-zA-Z]', input_omset_today)
    input_omset_today = input_omset_today.split(',')

    if len(cek_string) > 0 or input_omset_today[0] == '':
        return input_omset_harian(
                tanggal_hari_ini, jam, 'Maaf Bro harus angka tidak boleh teks')
    elif len(tanggal_hari_ini[0]) == 2  and '-' in tanggal_hari_ini[1]:
        jumlah_hari = tanggal_hari_ini[0][1] - tanggal_hari_ini[0][0]
        if len(input_omset_today) == jumlah_hari + 1:
            return [int(i) for i in input_omset_today]
        else:
            return input_omset_harian(
                tanggal_hari_ini, jam, 'Maaf input omset anda kurang atau kelebihan, silahkan lihat contoh format')
    elif len(tanggal_hari_ini[0]) > 1 and ',' in tanggal_hari_ini[1]:
        if len(tanggal_hari_ini) == len(input_omset_today):
            return [int(i) for i in input_omset_today]
        else:
            return input_omset_harian(
                tanggal_hari_ini, jam, 'Maaf input omset anda kurang atau kelebihan, silahkan lihat contoh format')
    else:
        if ',' in omset_tunggal:
            return input_omset_harian(
                tanggal_hari_ini, jam, 'Maaf ini input tunggal, tidak boleh ada koma')
        else:
            return [omset_tunggal]


