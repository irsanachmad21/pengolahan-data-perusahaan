import openpyxl
wb = openpyxl.load_workbook('Pagi_1_C.xlsx')
wb.active = 2 
sh = wb.active
print(sh)

#menghitung jumlah
print('saham penutupan')
jumlahJ = 0
for i in range(103,123):
    jumlahJ = jumlahJ + sh['J' + str(i)].value
print('jumlah saham penutupan = ',jumlahJ)
sh['B780'] = jumlahJ

print()

print('saham sebelumnya')
jumlahE = 0
for i in range(103,123):
    jumlahE = jumlahE + sh['E' + str(i)].value
print('jumlah saham sebelumnya = ',jumlahE)
sh['C780'] = jumlahE

print()

print('selisih saham')
for i in range(103,123):
    selisih = sh['J' + str(i)].value - sh['E' + str(i)].value
    print(f'nilai selisih pada kolom {i}= ', selisih)

#selisih keseluruhan
for i in range(103,123):
    selisih_keseluruhan = jumlahJ - jumlahE
print(selisih_keseluruhan)
sh['B781'] = selisih_keseluruhan

wb.save('hasilproses.xlsx')
