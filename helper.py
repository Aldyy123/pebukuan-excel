import locale
import re
from userController import tanggal_jam
from openpyxl import styles
import os

# Check String caracter when user entered string caracter
def cek_tanggal(tanggal, total_days):
    tanggal_urut = tanggal.split('-')
    tanggal_pilihan = tanggal.split(',')

    cek_string = re.findall('[a-zA-Z]', tanggal)


    if len(cek_string) > 0:
        return tanggal_jam('Masukan angka jangan string')
    elif '' in tanggal_urut or '' in tanggal_pilihan:
        return tanggal_jam('Masukan tanggal sesuai intruksi')
    elif(len(tanggal_urut) == 2 and '-' in tanggal):

        tanggal_urut = [int(i) for i in tanggal_urut]
        if tanggal_urut[0] < tanggal_urut[1] and tanggal_urut[1] < total_days:
            return tanggal_urut, '-'
        else:
            return tanggal_jam('Maaf tanggal tidak boleh kebalik atau lebih dari tanggal bulan ini')

    elif len(tanggal_pilihan) > 1 and ',' in tanggal:

        due_date = [int(i) < total_days for i in tanggal_pilihan]
        if all(due_date):
            tanggal_pilihan = [int(i) for i in tanggal_pilihan]
            filter_same_number = set(tanggal_pilihan)
            print('Tanggal yang sama akan mengambil 1 tanggalan yang sama dari beberapa tanggal contoh :', filter_same_number)
            return list(filter_same_number), ','
        else:
            return tanggal_jam('Maaf anda melebihi batas tanggalan bulan ini')
    elif int(tanggal) < total_days:
        return [int(tanggal)], ','
    else:
        return tanggal_jam('Maaf masukan angka yang valid')
        

def format_rupiah(duit):
    locale.setlocale(locale.LC_NUMERIC, 'IND')
    rupiah = locale.format("%.*f", (2, duit), True)
    return "Rp " + rupiah


# def check_tanggal_inptu_excel(tanggal):
#     if len(tanggal) == 2:
#         tanggal = tanggal[1] - tanggal[0]
#         for i in range(tanggal[0], tanggal[1])


def store_data_omset(dates, ws, omset, jam, separator):
   
    i = 0
    if len(dates) == 2 and '-' in separator:
        for tanggal in range(dates[0], dates[1] + 1):
            ws["A{}".format(tanggal + 3)].alignment = styles.Alignment(horizontal='left', vertical='center')
            ws["B{}".format(tanggal + 3)].alignment = styles.Alignment(horizontal='left', vertical='center')
            ws["C{}".format(tanggal + 3)].alignment = styles.Alignment(horizontal='left', vertical='center')
            
            ws['A{}'.format(tanggal + 3)] = tanggal
            ws['B{}'.format(tanggal + 3)] = omset[i]
            ws['C{}'.format(tanggal + 3)] = jam
            i = i + 1
    elif len(dates) > 0 and ',' in separator:
        for tanggal in dates:
            ws["A{}".format(tanggal + 3)].alignment = styles.Alignment(horizontal='left', vertical='center')
            ws["B{}".format(tanggal + 3)].alignment = styles.Alignment(horizontal='left', vertical='center')
            ws["C{}".format(tanggal + 3)].alignment = styles.Alignment(horizontal='left', vertical='center')
            
            ws['A{}'.format(tanggal + 3)] = tanggal
            ws['B{}'.format(tanggal + 3)] = omset[i]
            ws['C{}'.format(tanggal + 3)] = jam
            i = i + 1
    else:
        print('Eh ada yang error gaes nanti aja')

    return ws


def create_directory_check(file):
    if os.path.isdir('excel'):
        if os.path.isfile(f'./excel/{file}.xlsx'):
            return True, file
        else:
            return False, file
    else:
        os.makedirs('excel')
        return False, file