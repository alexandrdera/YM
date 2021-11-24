#!/usr/bin/env python
# coding: utf-8

# In[167]:


import pandas as pd
import time as t
import xlrd
from pathlib import Path
import shutil
from zipfile import ZipFile
import os


# In[168]:


ts1 = t.time()


# In[169]:


file_1c = t.strftime("%m%d") + "0900.xlsx"
file_path_1c = Path("D:/temp/_primary/", file_1c)
cumulative_file = Path("D:/Analytik/1_Основное планирование/", "01_ИСХОДНИК-НОВ (выгрузка 1С).xlsx")
file_path_out2 = Path("D:/temp/", "01_ИСХОДНИК-НОВ (выгрузка 1С)_ТЕСТ.xlsx")


# In[170]:


df_cf = pd.read_excel(cumulative_file, "TDSheet", usecols="A:O", index_col=False)


# In[171]:


print("Записей в накопетельном файле: ", len(df_cf))


# In[173]:


df_cf


# In[174]:


df_cf = df_cf.drop(df_cf.loc[df_cf['Категория клиента (св-во Контрагент)'] == "VIP Прямые продажи"].index)
df_cf = df_cf.drop(df_cf.loc[df_cf['Категория клиента (св-во Контрагент)'] == "Дистрибьютор"].index)


# In[175]:


print("Последняя запись в накопителе от ", df_cf["По дням"].max())


# In[176]:


file_1c_zip = file_1c.rpartition(".")[0]


# In[177]:


try:
    df_1с = pd.read_excel(file_path_1c, header = 7, usecols="B:P", index_col=False)
except KeyError as ke:
    # Создаем временную папку
    tmp_folder = Path('D:/temp/convert_wrong_excel/')
    os.makedirs(tmp_folder, exist_ok=True)

    # Распаковываем excel как zip в нашу временную папку
    with ZipFile(file_path_1c) as excel_container:
        excel_container.extractall(tmp_folder)

    # Переименовываем файл с неверным названием
    wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
    correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')

    os.rename(wrong_file_path, correct_file_path)

    # Запаковываем excel обратно в zip, перемещаем и переименовываем в исходный файл
    shutil.make_archive(file_1c_zip, 'zip', tmp_folder)
    shutil.move(os.path.realpath(file_1c_zip + ".zip"), str(file_path_1c))
    
    print("Файл перепакован из-за ошибки: "  + str(ke))

    df_1с = pd.read_excel(file_path_1c, header = 7, usecols="B:P", index_col=False)


# In[178]:


print("Из 1с загружено строк: ", len(df_1с))


# In[179]:


df_1с = df_1с.rename(columns={'Контрагент (категории)':'Категория клиента (св-во Контрагент)',
                     'Пользователь':'Основной менеджер покупателя'})


# In[180]:


df_1с["По дням"] = pd.to_datetime(df_1с["По дням"], format="%d.%m.%Y")


# In[181]:


df_vip1 = df_1с.loc[df_1с['Категория клиента (св-во Контрагент)'] == "VIP Прямые продажи"]
df_vip2 = df_1с.loc[df_1с['Категория клиента (св-во Контрагент)'] == "Дистрибьютор"]

df_vip = pd.concat([df_vip1, df_vip2])


# In[182]:


df_1с = df_1с.drop(df_1с.loc[df_1с['Категория клиента (св-во Контрагент)'] == "VIP Прямые продажи"].index)
df_1c_retail = df_1с.drop(df_1с.loc[df_1с['Категория клиента (св-во Контрагент)'] == "Дистрибьютор"].index)
print("Последняя запись в базе 1с от ", df_1c_retail["По дням"].max())


# In[183]:


df_cf_new = pd.concat([df_cf,
                       df_vip,
                       df_1c_retail.loc[df_1c_retail["По дням"] > df_cf["По дням"].max()]
                      ])


# In[184]:


df_cf_new


# In[185]:


df_cf_new = df_cf_new.dropna(subset = ["Количество (в ед. отчетов)"])


# In[186]:


df_cf_new.loc[df_cf_new["Контрагент"] == "Конечный оптовый покупатель", ["Грузополучатель"]] = ""


# In[187]:


df_obj = df_cf_new.select_dtypes(['object'])
df_cf_new[df_obj.columns] = df_obj.apply(lambda x: x.str.strip())


# In[192]:


import xlwings as xw

app_excel = xw.App(visible = True)

wb = xw.Book(cumulative_file)
ws = wb.sheets['TDSheet']
ws.range('A1').options(pd.DataFrame, index=False).value = df_cf_new

wb.api.RefreshAll()
wb.save()


# kill Excel process
'''
wb.close()
app_excel.kill()
del app_excel
'''


# In[193]:


#df_cf_new.to_excel(file_path_out2, sheet_name="TDSheet")


# In[190]:


print("Строк в новом файле: ", len(df_cf_new))


# In[191]:


print("Время выполнения ", t.time() - ts1)

