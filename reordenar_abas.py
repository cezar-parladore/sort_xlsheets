from threading import local
import xlwings as xw
import pandas as pd
from datetime import datetime


input_files = [r'C:\Users\cezar.parladore\Documentos\00_Arquivos_OFFline\2022-04-07\Dragagem\ASUP.xlsx',
               r'C:\Users\cezar.parladore\Documentos\00_Arquivos_OFFline\2022-04-07\Dragagem\SE.xlsx',
               r'C:\Users\cezar.parladore\Documentos\00_Arquivos_OFFline\2022-04-07\Dragagem\EF.xlsx']

def replace_pt_en(data:str):
    pt_to_en = {
        "Jan":"jan",
        "Feb":"fev",
        "Mar":"mar",
        "Apr":"abr",
        "Mai":"mai",
        "Jun":"jun",
        "Jul":"jul",
        "Aug":"ago",
        "Sep":"set",
        "Oct":"out",
        "Nov":"nov",
        "Dec":"dez"
    }
    mes = data.split(" ")[1]
    try:
        data = data.replace(mes, pt_to_en[mes])
        return data
    except:
        return data
    

for file in input_files:
    with xw.App() as app:
        book1 = app.books.open(file,read_only=False)
        abas_count = book1.sheets
        abas_nomes = [book1.sheets(i).name for i in abas_count]
        
        
        sorted_names = abas_nomes.copy()
        sorted_names.sort(key = lambda date: datetime.strptime(date, '%d %b %Y'))
        
    
        book2 = app.books.add()
        for sheet in sorted_names:
            i = sorted_names.index(sheet)
            book1.sheets(sheet).copy(after=book2.sheets[i], name=replace_pt_en(sheet))
        book2.sheets("Planilha1").delete()
        book2.save(f'{file.replace(".xlsx","_sorted.xlsx")}')