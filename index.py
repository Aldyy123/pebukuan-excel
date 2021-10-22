import userController
from colorama import Fore
chose_user_command = userController.commandUserList()
from excel import input_omset_harian, input_new_excel, update_excel

while True:
    if chose_user_command == 1:
        # Harian
        # Jika belum mengisi tanggal sebelumnya
        # Jika males ngisi tanggal default tanggal sekarang
        # Jika ingin meruntut dari tanggal ke tanggal berapa
        file = userController.insert_file()
        tanggal_jam = userController.tanggal_jam()
        omset_hari_ini = input_omset_harian(tanggal_jam[0], tanggal_jam[1])

        if file[0]:
            update_excel(file[1], tanggal_jam[0][0], tanggal_jam[1], omset_hari_ini, tanggal_jam[0][1])
        else:
            input_new_excel(file[1], tanggal_jam[0][0], tanggal_jam[1], omset_hari_ini, tanggal_jam[0][1])
        break
    elif chose_user_command == 2:
        print(2)
        break
    elif chose_user_command == 3:
        print(3)
        break
    else:
        print(Fore.RED + 'Maaf Input yang anda masukan salah'.upper() + Fore.RESET)
        chose_user_command = userController.commandUserList()
