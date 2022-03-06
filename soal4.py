from cProfile import label
import openpyxl
wb = openpyxl.load_workbook('Pagi_1_C.xlsx')
wb.active = 3
sh = wb.active
print(sh)

#membuat diagram bar
import numpy as np
x=np.zeros((200), 'U200')
y=np.zeros(200)

for i in range(123,143):
    x1 = sh['B' + str(i)].value
    x[i] = x1
    y1 = sh['L' + str(i)].value
    y[i] = y1
    print(x[i], y[i])

import matplotlib.pyplot as plt
plt.bar(x,y, label='blue bar', color='b')
plt.plot()
plt.xlabel('kode saham')
plt.ylabel('volume perusahaan')
plt.title('volume saham pada perusahaan')
plt.legend()
plt.show()