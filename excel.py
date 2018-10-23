from openpyxl import Workbook,load_workbook
from datetime import datetime
wb=load_workbook("kaoqin.xlsx")
ws = wb.active
first_column = ws['A']
mydict=dict()
for var in ws.iter_rows():
    mydict[var[0].value] = []

for var in ws.iter_rows():
    mydict[var[0].value].append(str(var[1].value))

duration=['2018-10-08','2018-10-09','2018-10-10','2018-10-11','2018-10-12','2018-10-13','2018-10-14','2018-10-15','2018-10-16','2018-10-17']
final=dict()

for var1 in mydict.keys():
    final[var1] = dict()
    for var2 in mydict[var1]:
        for var3 in duration:
            ATime, BTime=var2.split(' ')
            if var3==ATime:
                if final[var1].get(var3):
                    final[var1][var3].append(var2)
                else:
                    final[var1][var3]=[]
                    final[var1][var3].append(var2)
# print(final)
sheet=wb.create_sheet('统计结果')
wb.save(filename='kaoqin.xlsx')
temp = duration.copy()
temp.insert(0, '')
sheet.append(temp)
for var1 in final.keys():
    tempA=[var1,'','','','','','','','','','']
    for var2 in final[var1]:
          tempA.insert(duration.index(var2)+1,'上班时间：'+final[var1][var2][0]+'下班时间：'+final[var1][var2][-1])
    sheet.append(tempA)

wb.save('kaoqin.xlsx')
print("保存成功")
