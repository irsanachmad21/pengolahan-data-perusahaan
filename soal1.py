import openpyxl
wb = openpyxl.load_workbook('Pagi_1_C.xlsx')
wb.active = 0
sh = wb.active
print(sh)

#menghitung saham tertinggi
print('SAHAM TERTINGGI')
jumlah = 0
for i in range (163,183):
    jumlah = jumlah + sh['H' + str(i)].value
print('jumlah saham tertinggi = ',jumlah)
#rata-rata saham tertinggi
rata2 = jumlah/(183-163)
print(f'rata-rata saham tertinggi = {rata2:.2f}')
sh['B780'] = jumlah
sh['B781'] = rata2

print()

#menghitung saham terendah
print('SAHAM TERENDAH')
jumlah = 0
for i in range (163,183):
    jumlah = jumlah + sh['I' + str(i)].value
print('jumlah saham terendah = ',jumlah)
#rata-rata saham terendah
rata2 = jumlah/(183-163)
print(f'rata-rata saham terendah = {rata2:.2f}')
sh['C780'] = jumlah
sh['C781'] = rata2

print()

#menghitung saham penutupan
print('SAHAM PENUTUPAN')
jumlah = 0
for i in range (163,183):
    jumlah = jumlah + sh['J' + str(i)].value
print('jumlah saham penutupan = ',jumlah)
#rata-rata saham penutupan
rata2 = jumlah/(183-163)
print(f'rata-rata saham penutupan = {rata2:.2f}')
sh['D780'] = jumlah
sh['D781'] = rata2

wb.save('hasilproses.xlsx')