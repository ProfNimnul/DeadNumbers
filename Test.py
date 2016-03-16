# -*-coding: utf-8-*-
import os
import re
import xlwt
from dbfread import DBF
from easygui import *


def ToTextFile(list_of_deads):
    f=open("DeadPhones.txt",mode="w")
    geu_num=0
    for current in list_of_deads:
        c=int(current[0])
        if c>geu_num:
            geu_num = c
            f.write("ЖЭУ № "+str(geu_num)+"\n")
        f.write(current+"\n")
    f.close()
  
  

def ToExcelFile(list_of_deads):
    wb=xlwt.Workbook()
    account_col,phone_col=1,2

    cell_count=0
    ws=wb.add_sheet('Мертвые номера', cell_overwrite_ok=True)
    
    for current in list_of_deads:
        account,phone = current.split("\t")
        ws.write(cell_count,account_col,str(account))
        ws.write(cell_count,phone_col,str(phone))
        cell_count+=1
        try:
            wb.save(dirname+'\\deadnumbers.xls')
        except (SystemError,OSError):
            msgbox("Не могу создать файл Excel!", ok_button="Закрыть", title="Ошибка!")
            exit()            
        
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
    t = DBF(dirname+"\\"+name) # открыли файл базы данных

    for record in t: # проходим записи
        if len(record["Not call"]):
            phone, account=record["Telephone"],record["Account"]
            sets[i].add(account+"\t"+phone)
    i+=1


deads = set.intersection(*sets)
deads=list(deads)
deads.sort()

ToTextFile(deads)
ToExcelFile(deads)
exit()



