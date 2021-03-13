from TableClass import *
from openpyxl import load_workbook


wb = load_workbook("data.xlsx")
wb.active
ws = wb["Sheet1"]
print(Get_cell_value(ws,column=2,row=6))
Set_cell_value(ws,23,column=2,row=14)

lst1,lst2,lst3 = Create_table_from_definedSpace(wb)
# for each in lst1:
#     print(each)
# for each in lst2:
#     print(each)
# for each in lst3:
#     print(each)

# table = lst1[0]
# print(table.Get_title())
# table.Set_title("Duy is the best")
# print(table.Get_title())

# table = lst2[0]
# table.Set_key_name("Name:")
# table.Set_key_value("234")
# print(table.Get_key_name())
# print(table.Get_key_value())

table = lst3[0]

# # for each in lst:
# #     print(each)
data = {"id":6,"name":"Van","Score":4}
table.Add_new_row(data)
table.Change_row_value(1,data)

lst = table.Get_dataTable()
for each in lst:
    print(each)
# # dic = excelAPI.Load_definedname(wb)
# for key in dic.keys():
#    print(dic[key])




# print(Get_cell_value(ws,"B4"))
# print(Get_cell_value(ws,column=2,row=4))
# wb.save("data.xlsx")
