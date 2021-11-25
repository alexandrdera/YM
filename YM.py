import pandas as pd
from pathlib import Path
import os
import shutil
from zipfile import ZipFile

def read_excel_from_1c(file_path="", file_name="", header=0, usecols=None, index_col=False):
    """
    Фукнция конвертирует кривые excel файлы выгруженные из 1С записанные 
    с ошибкой: "There is no item named 'xl/sharedStrings.xml' in the archive"
    """
    try:
        return pd.read_excel(file_path, header = header, usecols=usecols, index_col=index_col)
    except KeyError as ke:
        # Создаем временную папку
        tmp_folder = Path('D:/temp/convert_wrong_excel/')
        os.makedirs(tmp_folder, exist_ok=True)

        # Распаковываем excel как zip в нашу временную папку
        with ZipFile(file_path) as excel_container:
            excel_container.extractall(tmp_folder)

        # Переименовываем файл с неверным названием
        wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
        correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')

        os.rename(wrong_file_path, correct_file_path)

        # Запаковываем excel обратно в zip, перемещаем и переименовываем в исходный файл
        file_1c_zip = file_name.rpartition(".")[0]
        shutil.make_archive(file_1c_zip, 'zip', tmp_folder)
        shutil.move(os.path.realpath(file_1c_zip + ".zip"), str(file_path))
        
        print("Файл перепакован из-за ошибки: "  + str(ke))

        return pd.read_excel(file_path, header = header, usecols=usecols, index_col=index_col)