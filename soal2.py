import openpyxl
wb = openpyxl.load_workbook('Pagi_1_C.xlsx')
wb.active = 1
sh = wb.active
print(sh)

#membuat diagram pie
import numpy as np
x=np.zeros((200), 'U200')
y=np.zeros(200)

for i in range(143,163):
    x1 = sh['B' + str(i)].value
    x[i] = x1
    y1 = sh['L' + str(i)].value
    y[i] = y1
    print(x[i], y[i])

import matplotlib.pyplot as plt
plt.pie(y, labels=x, autopct='%1.2f%%')
plt.title('volume perusahaan')
plt.show()
