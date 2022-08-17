import userController
from colorama import Fore
chose_user_command = userController.commandUserList()
from excel import user_omset_daily, input_new_excel, update_excel
from openpyxl.utils.exceptions import SheetTitleException

while True:
    if chose_user_command == 1:
        # Harian
        # Jika belum mengisi tanggal sebelumnya
        # Jika males ngisi tanggal default tanggal sekarang
        # Jika ingin meruntut dari tanggal ke tanggal berapa
        try:
            boolean_and_file = userController.insert_file()
            dates_omset = userController.input_dates()
            money_omset = user_omset_daily(dates_omset[0], dates_omset[1])
        except SheetTitleException:
            print("Error")

        if boolean_and_file[0]:
            update_excel(boolean_and_file[1], dates_omset[0][0], dates_omset[1], money_omset, dates_omset[0][1])
        else:
            input_new_excel(boolean_and_file[1], dates_omset[0][0], dates_omset[1], money_omset, dates_omset[0][1])
        break
    elif chose_user_command == 2:
        userController.insert_file()
        break
    elif chose_user_command == 3:
        print(3)
        break
    else:
        print(Fore.RED + 'Maaf Input yang anda masukan salah'.upper() + Fore.RESET)
        chose_user_command = userController.commandUserList()
