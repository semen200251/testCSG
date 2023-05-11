import logging
import os
import pandas as pd
import pythoncom
import win32com.client as win32
import time

def get_project(path):
    """Открывает файл проекта и возвращает объект проекта"""
    if not os.path.isabs(path):
        logging.warning('%s: Путь до файла проекта не абсолютный', get_project.__name__)
    logging.info('%s: Пытаемся открыть файл проекта', get_project.__name__)
    try:
        msp = win32.Dispatch("MSProject.Application", pythoncom.CoInitialize())
        _abs_path = os.path.abspath(path)
        print(_abs_path)
        msp.FileOpen(_abs_path)
        project = msp.ActiveProject
    except Exception:
        logging.error('%s: Файл проекта не смог открыться', get_project.__name__)
        raise Exception('Не получилось открыть файл проекта')
    logging.info('%s: Файл проекта успешно открылся', get_project.__name__)
    return project, msp


def fill_dataframe(project, columns):
    """Заполняет DataFrame значениями из project"""
    logging.info('%s: Создаем DataFrame из столбцов объекта проекта', fill_dataframe.__name__)
    if not project:
        logging.error('%s: Не удалось получить объект проекта', fill_dataframe.__name__)
        raise Exception("Объект проекта пустой")
    if not columns:
        logging.error('%s: Ключевые столбцы не заданы', fill_dataframe.__name__)
        raise Exception("Ключевые столбцы не заданы")
    task_collection = project.Tasks
    data = pd.DataFrame(columns=columns)
    try:
        for i in range(0,10):
            for t in task_collection:
                data.loc[len(data.index)] = [t.Name, t.Name, t.Name, t.Name, t.Name, t.Name, t.Name, t.Name, t.Name, t.Name]
    except Exception:
        logging.error('%s: Не получилось создать DataFrame из столбцов объекта проекта', fill_dataframe.__name__)
        raise Exception('Не получилось создать DataFrame из проекта')
    logging.info('%s: DataFrame из столбцов объекта проекта успешно создан', fill_dataframe.__name__)
    return data

if __name__ == "__main__":
    start = time.time()
    columns = [f"Имя{i}" for i in range(0,10)]
    for j in range(0, 10):
        project, mpp = get_project(r"C:\Users\semenhomec\PycharmProjects\pythonProject1\051-2000260_2022_нг_ф_(06.04).mpp")

        data = fill_dataframe(project, columns)
        data.to_excel(f"example{j}.xlsx")
        mpp.Quit()
        print(j, time.time() - start)
    end = time.time() - start
    print(end)



