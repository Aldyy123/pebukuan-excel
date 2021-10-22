import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
from matplotlib import style
from openpyxl import load_workbook
import os


if os.path.isfile('excel/apa.xlsx'):
    print('Ada')
else:
    print('Gak adaa')

# book = load_workbook('baru.xlsx')
# sheet = book['Oktober']

# a1 = sheet['B4']
# a2 = sheet['A4']
# a3 = sheet.cell(row=3, column=1)

# print(a1.value)
# print(a2.value) 
# print(a3.value)

# book.save('baru.xlsx')

# dfs = pd.read_excel('baru.xlsx', sheet_name='Oktober')

# print(dfs)

# style.use('ggplot')

# x = [0, 1, 2, 3, 4, 8]
# y = [24.27, 23.18, 22.39, 8.41, 7.19, 6.62]

# fig, ax = plt.subplots()

# ax.bar(x, y, align='center')

# ax.set_title('Omset bulan Oktober')
# ax.set_ylabel('Omset')
# ax.set_xlabel('Tanggal')

# ax.set_xticks(x)
# ax.set_xticklabels(("Python", "JavaScript", "Java", "C#", "PHP", "C++"))

# plt.show()
