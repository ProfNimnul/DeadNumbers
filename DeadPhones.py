# -*-coding: utf-8-*-
import os
import re
import xlwt
from dbfread import DBF
from easygui import *


def ToTextFile(data_to_write): f.write(data_to_write + "\n")
  
def ToExcelFile(data_to_write, cell):
    account_col=0
    phone_col=1
    account, phone = data_to_write.split("\t")
    ws.write(cell, account_col, str(account))
    ws.write(cell, phone_col, str(phone))
    

dbflist = set()

dirname = diropenbox("Укажите каталог с файлами протоколов", "")
if not dirname:
    dirname = os.curdir
# print(dirname)

# l = os.listdir("d:/py")
l = os.listdir(dirname)

for x in l:
    x = x.lower()
    if x.endswith('.dbf'):
        rezult = re.match("^obz(von)?(0[1-9]|1[0-2])[1-9][5-9]\.", x, flags=re.IGNORECASE)
        if rezult != None:
            dbflist.add(x)

lendbflist = len(dbflist)

if not lendbflist:
    msgbox("Отсутствуют файлы dbf", ok_button="Закрыть", title="Отсутствуют файлы!")
    exit()

sets = [set()] * lendbflist  # нашли пересечение множеств

i = 0  # cчетчик множеств в метамножестве

for name in dbflist:
    #   t = DBF("d:/py/"+name)
    t = DBF(dirname + "\\" + name)  # открыли файл базы данных

    for record in t:  # проходим записи
        if len(record["Not call"]):
            phone, account = record["Telephone"], record["Account"]
            sets[i].add(account + "\t" + phone)
    i += 1

deads = set.intersection(*sets)
deads = list(deads)
deads.sort()

# для вывода в текст    
try:
    f = open("DeadPhones.txt", mode="w")
except (SystemError, OSError):
    msgbox("Не могу создать файл TXT!", ok_button="Закрыть", title="Ошибка!")

# Для вывода в Эксель 
wb = xlwt.Workbook()
ws = wb.add_sheet('Мертвые номера', cell_overwrite_ok=True)

geu_num=0
cell_count = 0

for one_abonent in deads:
    new_geu = int(one_abonent[0])
    if new_geu > geu_num:
        geu_num=new_geu    
        data_to_write="ЖЭУ № " + str(geu_num)+"\t"+""
    else:
        data_to_write=one_abonent    
    
    ToTextFile(data_to_write)
    ToExcelFile(data_to_write, cell_count )
    cell_count += 1
f.close()

try:
    wb.save(dirname + '\\deadnumbers.xls')
except (SystemError, OSError):
    msgbox("Не могу создать файл Excel!", ok_button="Закрыть", title="Ошибка!")

exit()
