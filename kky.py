# IMPORTS
import datetime
import decimal
import os
import re
# import shutil
import time

import numpy as np
import openpyxl
import pandas as pd
# import sidetable

import функции

pd.set_option("display.max_rows", 1500)
pd.set_option("display.max_columns", 100)
pd.set_option("max_colwidth", 30)
pd.set_option("expand_frame_repr", False)
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# global variables
USERPROFILE = os.environ["USERPROFILE"]

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty lists

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty dictionaries

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty dataframes
df_реализация_fin = pd.DataFrame()
findf = pd.DataFrame()
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# default lists
OP_list = ["Агрин", "Графовская", "Коренская", "Муромская", "Нежегольская", "Полянская", "Томаровская", "Валуйская", "Рождественская"]
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# default dictionaries

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# prompts for user input
prompt1 = "\nреализация?: "
prompt2 = "\nзаписать в отчет напрямую?: "
prompt3 = "\nномер строки для записи в отчет напрямую (нумерация excel) - бройлеры?: "
prompt4 = "\nномер строки для записи в отчет напрямую (нумерация excel) - старка?: "

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# user inputs
inp1 = input(prompt1)
# inp2 = input(prompt2)
inp2 = "нет"
if inp2 == "да" or inp2 == "yes" or inp2 == "y":
    inp3 = input(prompt3)
    inp3 = int(inp3) - 1
    inp4 = input(prompt4)
    inp4 = int(inp4) - 1

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# file paths
path_кку = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\оу контрольная карта убой\\"
path_реализация = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\реализация\\"
listoffiles_кку = os.listdir(path_кку)
listoffiles_реализация = os.listdir(path_реализация)
# file names
filename0a = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\_промежуточный файл df_впк.xlsx"
filename0b = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\_промежуточный файл df_бройлеры.xlsx"
filename0c = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\_промежуточный файл df_старка.xlsx"
filename1 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\Динамика развития птицы Белгород.xlsx"
filename2 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\сп-51\\сп-51.xlsx"
filename3 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\Накопительный 2023 - Белгород.xlsx"
filename4 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\Накопительный (старка) Белгород 2023.xlsx"
filename5 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\время поднятия кормушки\\время поднятия кормушки.xlsx"

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# ВРЕМЯ ПОДНЯТИЯ КОРМУШКИ---------------------------------------------------------------------
df_from_excel = pd.read_excel(
    filename5,
    sheet_name="TDSheet",
    # index_col=0,
    # engine = "openpyxl",
    header=0,
    usecols = "H,J,K,L,O,Q,U",
    )
df_from_excel = df_from_excel.loc[(df_from_excel["Вид выбытия"].str.contains("Основная")) | (df_from_excel["Вид выбытия"].str.contains("Разрежение"))]
df_from_excel = df_from_excel.dropna(subset=["Время поднятия кормушки"])
# print(df_from_excel.head())
# exit()
#
df_from_excel.reset_index(inplace = True)
df_from_excel = df_from_excel.drop(["index"], axis = 1)
df_from_excel.loc[df_from_excel["Корпус"].str.contains(" пл."), ["площ"]] = df_from_excel["Корпус"]
df_from_excel.loc[df_from_excel["площ"].str.contains("Агрин"), ["площ"]] = "Агрин"
df_from_excel.loc[df_from_excel["площ"].str.contains("Коренская"), ["площ"]] = "Коренское"
df_from_excel.loc[df_from_excel["площ"].str.contains("Графовская"), ["площ"]] = "Графовское"
df_from_excel.loc[df_from_excel["площ"].str.contains("Полянская"), ["площ"]] = "Полянское"
df_from_excel.loc[df_from_excel["площ"].str.contains("Томаровская"), ["площ"]] = "Томаровское"
df_from_excel.loc[(df_from_excel["площ"].str.contains("Муромская")==True) & (df_from_excel["площ"].str.contains(" РМ")==False) & (df_from_excel["площ"].str.contains(" РС")==False), ["площ"]] = "Муромское"
df_from_excel.loc[df_from_excel["площ"].str.contains("Нежегольская"), ["площ"]] = "Нежегольское"
df_from_excel.loc[df_from_excel["площ"].str.contains("Валуйская"), ["площ"]] = "Валуйское"
df_from_excel.loc[df_from_excel["площ"].str.contains("Рождественская"), ["площ"]] = "Рождественское"
#
df_from_excel.loc[df_from_excel["Корпус"].str.contains("корпус"), ["корп"]] = df_from_excel["Корпус"].str[0:3]
df_from_excel.loc[df_from_excel["корп"].str.contains(" к", na=False), ["корп"]] = df_from_excel["корп"].str[0:2]
df_from_excel["корп"] = "_" + df_from_excel["корп"].astype(str) + "_"
df_from_excel["корп"] = df_from_excel["корп"].apply(lambda x: x.replace(" ","")) # здесь не пробел, а специальный символ из 1С
df_from_excel["корп"] = df_from_excel["корп"].apply(lambda x: x.replace("_",""))
df_from_excel.loc[df_from_excel["площ"].str.contains("Муромское"), ["корп"]] = df_from_excel["корп"].astype(str).str[:1] + "." +df_from_excel["корп"].astype(str).str[1:]
df_from_excel.loc[df_from_excel["площ"].str.contains("Муромск"), ["корп"]] = df_from_excel["корп"].apply(lambda x: x.replace(".",","))
df_from_excel["корп"] = df_from_excel["корп"].apply(lambda x: float(x) if str(x).isdigit() else x)
#
df_from_excel = df_from_excel.drop(["Корпус"], axis = 1)
print("\ndf_from_excel")
print(df_from_excel)
#
df_pivot = pd.pivot_table(
    df_from_excel,
    # index=["Дата и время посадки/выбытия", "площ", "Вид выбытия",  "корп", "Время поднятия кормушки"],
    index=["Дата и время посадки/выбытия", "площ", "корп", "Время поднятия кормушки"],
    columns=["Причина движения"],
    # values=["Головы"],
    values=["Головы", "Вес, кг"],
    # aggfunc="mean",
    fill_value=0,
    )
df_pivot.columns = ['_'.join(col) for col in df_pivot.columns.values]
df_pivot.reset_index(inplace = True)
df_pivot = df_pivot.rename(columns={
    "Головы_Внешняя реализация": "Р-ция голов",
    "Головы_На мясо": "На мясо голов",
    "Головы_Падёж в пути": "Падеж голов",
    "Головы_Сдано на комбинат": "Живок голов",
    #
    "Вес, кг_Внешняя реализация": "Р-ция вес",
    "Вес, кг_На мясо": "На мясо вес",
    "Вес, кг_Падёж в пути": "Падеж вес",
    "Вес, кг_Сдано на комбинат": "Живок вес",
    })
try:
    df_pivot = df_pivot.drop(["Р-ция голов"], axis = 1)
    df_pivot = df_pivot.drop(["На мясо голов"], axis = 1)
    df_pivot = df_pivot.drop(["Р-ция вес"], axis = 1)
    df_pivot = df_pivot.drop(["На мясо вес"], axis = 1)
except KeyError:
    pass
df_pivot["Живок голов"] = df_pivot["Живок голов"] + df_pivot["Падеж голов"]
df_pivot["Живок вес"] = df_pivot["Живок вес"] + df_pivot["Падеж вес"]
df_pivot["Дата и время посадки/выбытия"] = pd.to_datetime(df_pivot["Дата и время посадки/выбытия"], dayfirst=True)
df_pivot = df_pivot.sort_values(by=["Дата и время посадки/выбытия", "площ"], ascending=True)
df_pivot = df_pivot[[
    "площ",
    "корп",
    "Живок голов",
    "Живок вес",
    "Падеж голов",
    "Падеж вес",
    "Время поднятия кормушки",
    "Дата и время посадки/выбытия"
    ]]
"""
# df_pivot["Живок голов"] = df_pivot["Живок голов"].astype(float)
df_pivot["Живок вес"] = df_pivot["Живок вес"].astype(float)
# df_pivot["Падеж голов"] = df_pivot["Падеж голов"].astype(float)
df_pivot["Падеж вес"] = df_pivot["Падеж вес"].astype(float)

df_pivot["Живок голов"] = pd.to_numeric(df_pivot["Живок голов"], errors="coerce")
df_pivot["Живок вес"] = pd.to_numeric(df_pivot["Живок вес"], errors="coerce")
df_pivot["Падеж голов"] = pd.to_numeric(df_pivot["Падеж голов"], errors="coerce")
df_pivot["Падеж вес"] = pd.to_numeric(df_pivot["Падеж вес"], errors="coerce")
df_pivot["Живок вес"] = df_pivot["Живок вес"].apply(lambda x: decimal.Decimal(x))
df_pivot["Живок вес"] = df_pivot["Живок вес"].apply(lambda x: x.quantize(decimal.Decimal("0.00")))
df_pivot["Падеж голов"] = df_pivot["Падеж голов"].apply(lambda x: decimal.Decimal(x))
df_pivot["Падеж голов"] = df_pivot["Падеж голов"].apply(lambda x: x.quantize(decimal.Decimal("0.0")))
"""
# df_pivot["keycol"] = df_pivot["площ"] + df_pivot["корп"].astype(str) + df_pivot["Живок голов"].astype(str) + df_pivot["Живок вес"].astype(str) + df_pivot["Падеж голов"].astype(str)+ df_pivot["Падеж вес"].astype(str)
print("\ndf_pivot")
print(df_pivot)
# exit()
функции.pd_toexcel(
            pd,
            #
            filename = filename0a,
            разновидность = "Лист1",
            df_для_записи = df_pivot,
            header_pd = "True",
            rowtostartin_pd = 0,
            coltostartin_pd = 0,
        )

# РЕАЛИЗАЦИЯ---------------------------------------------------------------------
if inp1 == "да" or inp1 == "yes" or inp1 == "y":
    # сп-51---------------------------------------------------------------------
    df_сп_51 = pd.read_excel(
        filename2,
        sheet_name="Лист1",
        index_col=0,
        engine = "openpyxl",
        header=15,
        usecols = "A,Q,R",
        )
    df_сп_51.reset_index(inplace = True)
    df_сп_51 = df_сп_51.rename(columns={"Пол": "площ", "голов.5": "гол", "масса, кг.6": "вес"})
    df_сп_51 = df_сп_51.dropna(subset=["гол"])
    df_сп_51 = df_сп_51.drop(df_сп_51.loc[df_сп_51["площ"].str.contains("Площадка")==False].index)
    print("\ndf_сп_51")
    print(df_сп_51)
    inp0 = input("Продолжить?:")
    if inp0 == "нет":
        exit()

    # реализация---------------------------------------------------------------------
    for i in listoffiles_реализация:
        regexpr = re.compile(r"(\d{2,4})+(.)")
        датасдачи2 = ""
        for gr in regexpr.findall(i):
            датасдачи2 = датасдачи2 + gr[0]
            датасдачи2 = датасдачи2 + gr[1]
        датасдачи2 = датасдачи2[0:-1]
        print("\nдата_сдачи_реализация")
        print(датасдачи2)
        wb = openpyxl.load_workbook(path_реализация + i)
        ws = wb["Лист1"]
        rowmax = ws.max_row + 1
        # print(rowmax)
        площадка = str(ws.cell(row = 3, column = 1).value)
        площадка = площадка.replace("Суточное движение поголовья Площадка ","")
        df_реализация = pd.read_excel(
            path_реализация + i,
            sheet_name="Лист1",
            index_col=0,
            engine = "openpyxl",
            header=7,
            usecols = "A,M,N",
            )
        df_реализация.reset_index(inplace = True)
        df_реализация = df_реализация.rename(columns={"index": "корп", "гол..5": "Живок голов", "вес.5": "Живок вес"})
        df_реализация = df_реализация.dropna(subset=["Живок голов"])
        df_реализация = df_реализация.drop(df_реализация.loc[df_реализация["корп"].str.contains("Итого", na=False)].index)
        df_реализация = df_реализация.drop(df_реализация.loc[df_реализация["корп"].str.contains("Всего", na=False)].index)
        df_реализация["Живок вес"]=df_реализация["Живок вес"].astype(str)
        df_реализация["Живок вес"] = df_реализация["Живок вес"].str.replace(" ","")
        df_реализация["Живок вес"] = df_реализация["Живок вес"].str.replace(",",".")
        df_реализация["Живок вес"] = pd.to_numeric(df_реализация["Живок вес"], errors="coerce")
        # 
        df_реализация["Живок голов"]=df_реализация["Живок голов"].astype(str)
        df_реализация["Живок голов"] = df_реализация["Живок голов"].str.replace(" ","")
        df_реализация["Живок голов"] = df_реализация["Живок голов"].str.replace(",",".")
        df_реализация["Живок голов"] = pd.to_numeric(df_реализация["Живок голов"], errors="coerce")
        # 
        df_реализация = df_реализация.groupby(["корп"], as_index=False).agg({"Живок голов": "sum", "Живок вес": "sum"})
        df_реализация["старка"] = площадка
        df_реализация["направ"] = "Центр"
        df_реализация["комб"] = "Реализация"
        df_реализация["Падеж голов"] = 0
        df_реализация["Падеж вес"] = 0
        df_реализация["дата.сдачи"] = датасдачи2
        df_реализация = функции.pd_movecol(
            df_реализация,
            cols_to_move=["старка", "направ", "комб"],
            ref_col="корп",
            place="Before"
            )
        df_реализация = функции.pd_movecol(
            df_реализация,
            cols_to_move=["дата.сдачи"],
            ref_col="корп",
            place="After"
            )
        df_реализация_fin = pd.concat([df_реализация_fin, df_реализация], ignore_index = True)
        print("\ndf_реализация")
        print(df_реализация)
        # print(df_реализация.dtypes)
    # exit()

# ДИНАМИКА РАЗВИТИЯ ПТИЦЫ---------------------------------------------------------------------
df_динамика = pd.read_excel(
    filename1,
    sheet_name="Лист1",
    index_col=0,
    engine = "openpyxl",
    header=1,
    usecols = "B,X",
    )
df_динамика.reset_index(inplace = True)
df_динамика["корп"] = df_динамика["№ корпуса"]
#
df_динамика["корп"] = "_" + df_динамика["корп"].astype(str) + "_"
# df_динамика["корп"] = df_динамика["корп"].map(lambda x: x.rstrip(" ")) # здесь не пробел, а специальный символ из 1С
# df_динамика["корп"] = df_динамика["корп"].str.replace(" ","") # здесь не пробел, а специальный символ из 1С
df_динамика["корп"] = df_динамика["корп"].apply(lambda x: x.replace(" ","")) # здесь не пробел, а специальный символ из 1С
df_динамика["корп"] = df_динамика["корп"].apply(lambda x: x.replace("_",""))
#
df_динамика["Дата посадки"] = pd.to_numeric(df_динамика["Дата посадки"], errors="coerce")
df_динамика["Дата посадки"] = pd.to_datetime(df_динамика["Дата посадки"], dayfirst=True, unit="D", origin="1899-12-30")
df_динамика.loc[df_динамика["№ корпуса"].apply(lambda x: x not in ["Агрин", "Графовское", "Коренское", "Муромское", "Нежегольское", "Полянское", "Томаровское", "Валуйское", "Рождественское"]), ["№ корпуса"]] = np.nan
df_динамика["№ корпуса"] = df_динамика["№ корпуса"].fillna(method="ffill")
df_динамика = df_динамика.rename(columns={"№ корпуса": "площ"})
df_динамика = df_динамика.dropna(subset=["Дата посадки"])
# print("\ndf_динамика")
# print(df_динамика)
# print(df_динамика.dtypes)
# exit()

# ККУ---------------------------------------------------------------------
for i in listoffiles_кку:
    # using regex to extract date from filename------------------------------
    regexpr = re.compile(r"(\d{1,4})+(.)")
    датасдачи = ""
    for gr in regexpr.findall(i):
        датасдачи = датасдачи + gr[0]
        датасдачи = датасдачи + gr[1]
    датасдачи = датасдачи[0:-1]
    функции.print_line("hyphens")
    print("\nдата_сдачи_кку")
    print(датасдачи)
    today = datetime.datetime.strptime(датасдачи, "%d.%m.%Y")
    weeknum = datetime.datetime.date(today).isocalendar().week
    # exit()

    # copying a row of colnames to be used as pandas header--------------------------
    wb = openpyxl.load_workbook(path_кку + i)
    ws = wb["Лист1"]
    # rowmax = ws.max_row + 1
    for b in range(7, 18):
        searchcell = str(ws.cell(row = 4, column = b).value)
        # print(searchcell)
        # target_cell = ws.cell(row = 9, column = b).value
        # target_cell = searchcell # this doesn"t work
        ws.cell(row = 9, column = b).value = searchcell # this is the correct way
    функции.wb_save_openpyxl(wb, path_кку + i)
    # exit()

    # ---------------------------------------------------------------------
    df_from_excel = pd.read_excel(
        path_кку + i,
        sheet_name="Лист1",
        index_col=0,
        engine = "openpyxl",
        header=8,
        usecols = "A,B,F,H,I,K,L,M",
        )
    df_from_excel.reset_index(inplace = True)
    df_from_excel = df_from_excel.rename(columns={
        "ТТН.Время прибытия": "прибытие",
        "Выгрузка окончание (по счетчику)": "выгрузка",
        "Содержимое зобов и ЖКТ, кг": "жкт",
        })
    df_from_excel["площ"] = np.nan
    df_from_excel["сдача"] = np.nan
    df_from_excel["корп"] = np.nan
    df_from_excel["старка"] = np.nan
    df_from_excel["нед"] = weeknum
    df_from_excel["дата.сдачи"] = датасдачи
    df_from_excel = функции.pd_movecol(
        df_from_excel,
        cols_to_move=["площ", "сдача", "корп", "старка", "нед", "дата.сдачи"],
        ref_col="index",
        place="After"
        )
    df_from_excel["index"] = df_from_excel["index"].fillna("XXX")
    #
    # df_from_excel.loc[df_from_excel["index"].str.contains("СтБр"), ["площ"]] = df_from_excel["index"].map(lambda x: x.rstrip(" СтБр"))
    df_from_excel.loc[df_from_excel["index"].apply(lambda x: x in OP_list), ["площ"]] = df_from_excel["index"]
    df_from_excel.loc[(df_from_excel["index"].str.contains("РС")==True) & (df_from_excel["index"].str.contains("корпус")==False), ["площ"]] = df_from_excel["index"]
    df_from_excel.loc[(df_from_excel["index"].str.contains("РМ")) & (df_from_excel["index"].str.contains("корпус")==False), ["площ"]] = df_from_excel["index"]
    df_from_excel["площ"] = df_from_excel["площ"].fillna(method="ffill")
    #
    df_from_excel.loc[df_from_excel["index"].str.contains("Основная"), ["сдача"]] = "Основная"
    df_from_excel.loc[df_from_excel["index"].str.contains("Разрежение"), ["сдача"]] = "Разрежение"
    df_from_excel["сдача"] = df_from_excel["сдача"].fillna(method="ffill")
    #
    df_from_excel.loc[df_from_excel["index"].str.contains("корпус"), ["корп"]] = df_from_excel["index"].str[0:3]
    df_from_excel.loc[df_from_excel["корп"].str.contains(" к", na=False), ["корп"]] = df_from_excel["корп"].str[0:2]
    df_from_excel["корп"] = df_from_excel["корп"].fillna(method="ffill")
    df_from_excel["корп"] = df_from_excel["корп"].fillna("XXX")
    #
    df_from_excel.loc[(df_from_excel["площ"].str.contains("РС")==False) & (df_from_excel["площ"].str.contains("РМ")==False), ["старка"]] = "нет"
    df_from_excel.loc[(df_from_excel["площ"].str.contains("РС")==True) | (df_from_excel["площ"].str.contains("РМ")==True), ["старка"]] = df_from_excel["площ"]
    df_from_excel["старка"] = df_from_excel["старка"].fillna(method="ffill")
    df_from_excel["старка"] = df_from_excel["старка"].fillna("нет")
    #
    df_from_excel = df_from_excel.dropna(subset=["выгрузка"])
    df_from_excel = df_from_excel.drop(["index"], axis = 1)
    df_from_excel["Падеж голов"] = df_from_excel["Падеж голов"].fillna(0)
    df_from_excel["Падеж вес"] = df_from_excel["Падеж вес"].fillna(0)
    df_from_excel["Живок голов"] = df_from_excel["Живок голов"].str.replace(" ","")
    df_from_excel["Живок вес"] = df_from_excel["Живок вес"].str.replace(" ","")
    df_from_excel["Живок вес"] = df_from_excel["Живок вес"].str.replace(",",".")
    df_from_excel["Падеж вес"] = df_from_excel["Падеж вес"].str.replace(" ","")
    df_from_excel["Падеж вес"] = df_from_excel["Падеж вес"].str.replace(",",".")
    df_from_excel["Живок голов"] = pd.to_numeric(df_from_excel["Живок голов"], errors="coerce")
    df_from_excel["Живок вес"] = pd.to_numeric(df_from_excel["Живок вес"], errors="coerce")
    df_from_excel["Падеж голов"] = pd.to_numeric(df_from_excel["Падеж голов"], errors="coerce")
    df_from_excel["Падеж вес"] = pd.to_numeric(df_from_excel["Падеж вес"], errors="coerce")
    df_from_excel["Падеж вес"] = df_from_excel["Падеж вес"].fillna(0)
    df_from_excel["площ"] = df_from_excel["площ"].astype(str)
    #
    # print("\ndf_from_excel")
    # print(df_from_excel)
    # exit()
    df_from_excel["жкт"] = df_from_excel["жкт"].fillna("")
    df_from_excel["жкт"] = df_from_excel["жкт"].str.replace(",",".")
    df_from_excel["жкт"] = pd.to_numeric(df_from_excel["жкт"], errors="coerce")
    df_from_excel["жкт"] = df_from_excel["жкт"].fillna(0)
    df_from_excel["Живок вес"] = df_from_excel["Живок вес"] - df_from_excel["жкт"]
    df_from_excel = df_from_excel.drop(["жкт"], axis = 1)
    #
    df_from_excel.loc[df_from_excel["площ"].str.contains("Коренская"), ["площ"]] = "Коренское"
    df_from_excel.loc[df_from_excel["площ"].str.contains("Графовская"), ["площ"]] = "Графовское"
    df_from_excel.loc[df_from_excel["площ"].str.contains("Полянская"), ["площ"]] = "Полянское"
    df_from_excel.loc[df_from_excel["площ"].str.contains("Томаровская"), ["площ"]] = "Томаровское"
    df_from_excel.loc[(df_from_excel["площ"].str.contains("Муромская")==True) & (df_from_excel["площ"].str.contains(" РМ")==False) & (df_from_excel["площ"].str.contains(" РС")==False), ["площ"]] = "Муромское"
    df_from_excel.loc[df_from_excel["площ"].str.contains("Нежегольская"), ["площ"]] = "Нежегольское"
    df_from_excel.loc[df_from_excel["площ"].str.contains("Валуйская"), ["площ"]] = "Валуйское"
    df_from_excel.loc[df_from_excel["площ"].str.contains("Рождественская"), ["площ"]] = "Рождественское"
    #
    df_from_excel["корп"] = "_" + df_from_excel["корп"].astype(str) + "_"
    # df_from_excel["корп"] = df_from_excel["корп"].map(lambda x: x.rstrip(" ")) # здесь не пробел, а специальный символ из 1С
    # df_from_excel["корп"] = df_from_excel["корп"].str.replace(" ","") # здесь не пробел, а специальный символ из 1С
    df_from_excel["корп"] = df_from_excel["корп"].apply(lambda x: x.replace(" ","")) # здесь не пробел, а специальный символ из 1С
    df_from_excel["корп"] = df_from_excel["корп"].apply(lambda x: x.replace("_",""))
    df_from_excel.loc[(df_from_excel["площ"].str.contains("Муромское")) & (df_from_excel["старка"] == "нет"), ["корп"]] = df_from_excel["корп"].astype(str).str[:1] + "." +df_from_excel["корп"].astype(str).str[1:]
    #
    # print("\ndf_from_excel")
    # print(df_from_excel)
    # print(df_from_excel.dtypes)
    # exit()
    df_from_excel = pd.merge(df_from_excel, df_динамика, how = "left", on = ["площ", "корп"])
    df_from_excel.loc[df_from_excel["площ"].str.contains("Муромск"), ["корп"]] = df_from_excel["корп"].apply(lambda x: x.replace(".",","))
    # df_from_excel["корп"] = pd.to_numeric(df_from_excel["корп"], errors="ignore")
    df_from_excel["корп"] = df_from_excel["корп"].apply(lambda x: float(x) if str(x).isdigit() else x)
    df_from_excel = функции.pd_movecol(
        df_from_excel,
        cols_to_move=["Дата посадки"],
        ref_col="корп",
        place="After"
        )
    print("\nЦБ")
    print(df_from_excel)
    # exit()

    # findf-----------------------------------------------------------------------------------------------------------------------------------------------------------
    findf = pd.concat([findf, df_from_excel], ignore_index = True)
    findf["прибытие"] = pd.to_datetime(findf["прибытие"], dayfirst=True)
    findf["выгрузка"] = pd.to_datetime(findf["выгрузка"], dayfirst=True)

findf["дата.сдачи.dt"] = pd.to_datetime(findf["дата.сдачи"], dayfirst=True)
findf = findf.sort_values(by=["дата.сдачи.dt", "площ", "корп", "прибытие"], ascending=True)
findf = findf.drop(["дата.сдачи.dt"], axis = 1)
функции.print_line("hyphens")
print("\nИТОГО")
print(findf)
# print(findf.dtypes)
# exit()


# writing to excel-------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# БРОЙЛЕРЫ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
df_бройлеры = findf.copy(deep=True)
df_бройлеры = df_бройлеры.drop(df_бройлеры[(df_бройлеры["старка"] != "нет")].index)
df_бройлеры = df_бройлеры.drop(["старка"], axis = 1)
df_бройлеры["гол"] = ""
df_бройлеры["кг"] = ""
df_бройлеры["комб"] = "Шебекинский ПК"
df_бройлеры["перев"] = ""
df_бройлеры["вид"] = "белая"
df_бройлеры["маш"] = ""
df_бройлеры = функции.pd_movecol(
        df_бройлеры,
        cols_to_move=["гол", "кг", "комб", "перев"],
        ref_col="площ",
        place="After"
        )
df_бройлеры = функции.pd_movecol(
        df_бройлеры,
        cols_to_move=["вид"],
        ref_col="сдача",
        place="After"
        )
df_бройлеры = функции.pd_movecol(
        df_бройлеры,
        cols_to_move=["маш"],
        ref_col="корп",
        place="After"
        )
"""
# df_бройлеры = pd.merge(df_бройлеры, df_pivot, how = "left", on = ["площ", "корп", "Живок голов", "Живок вес", "Падеж голов", "Падеж вес"])
df_pivot["keycol"] = df_pivot["площ"] + df_pivot["корп"].astype(str) + df_pivot["Живок голов"].astype(str) + df_pivot["Живок вес"].astype(str) + df_pivot["Падеж голов"].astype(str)+ df_pivot["Падеж вес"].astype(str)
df_бройлеры["keycol"] = df_бройлеры["площ"] + df_бройлеры["корп"].astype(str) + df_бройлеры["Живок голов"].astype(str) + df_бройлеры["Живок вес"].astype(str) + df_бройлеры["Падеж голов"].astype(str)+ df_бройлеры["Падеж вес"].astype(str)
df_бройлеры = pd.merge(df_бройлеры, df_pivot, how = "left", on = ["keycol"])
# df_бройлеры = df_бройлеры.drop(["keycol"], axis = 1)
"""
# print("\ndf_бройлеры")
# print(df_бройлеры)

if inp2 == "нет" or inp2 == "no" or inp2 == "n":
    функции.pd_toexcel(
                pd,
                #
                filename = filename0b,
                разновидность = "Лист1",
                df_для_записи = df_бройлеры,
                header_pd = "True",
                rowtostartin_pd = 0,
                coltostartin_pd = 0,
            )
if inp2 == "да" or inp2 == "yes" or inp2 == "y":
    start = time.time()
    df_бройлеры_part1 = df_бройлеры.copy(deep=True)
    df_бройлеры_part1 = df_бройлеры_part1.drop(["нед", "дата.сдачи", "прибытие", "выгрузка", "Живок голов", "Живок вес", "Падеж голов", "Падеж вес"], axis = 1)
    # print("\ndf_бройлеры_part1")
    # print(df_бройлеры_part1)

    df_бройлеры_part2 = df_бройлеры.copy(deep=True)
    df_бройлеры_part2 = df_бройлеры_part2[["нед", "дата.сдачи", "прибытие", "выгрузка"]]
    # print("\ndf_бройлеры_part2")
    # print(df_бройлеры_part2)

    df_бройлеры_part3 = df_бройлеры.copy(deep=True)
    df_бройлеры_part3 = df_бройлеры_part3[["Живок голов", "Живок вес"]]
    # print("\ndf_бройлеры_part3")
    # print(df_бройлеры_part3)

    df_бройлеры_part4 = df_бройлеры.copy(deep=True)
    df_бройлеры_part4 = df_бройлеры_part4[["Падеж голов", "Падеж вес"]]
    # print("\ndf_бройлеры_part4")
    # print(df_бройлеры_part4)

    функции.df_to_excel_openpyxl(
        filename = filename3,
        разновидность = "Убой ШПК",
        df_для_записи = df_бройлеры_part1,
        rowtostartin_pd = inp3,
        coltostartin_pd = 1,
        всего_colnum_offset = 2,
        неприказ_belowtablenames_offset = 0,
        приказ_belowtablenames_offset = 0,
        clearing_marker = "не удалять",
        clearing_marker_col = 1,
        clearing_offset = 1,
        remove_borders = 0,
        change_alignment = 0,
        add_borders = 0,
        aggr_row = 0,
        font_change_scope = 0,
    )
    функции.df_to_excel_openpyxl(
        filename = filename3,
        разновидность = "Убой ШПК",
        df_для_записи = df_бройлеры_part2,
        rowtostartin_pd = inp3,
        coltostartin_pd = 13,
        всего_colnum_offset = 2,
        неприказ_belowtablenames_offset = 0,
        приказ_belowtablenames_offset = 0,
        clearing_marker = "не удалять",
        clearing_marker_col = 1,
        clearing_offset = 1,
        remove_borders = 0,
        change_alignment = 0,
        add_borders = 0,
        aggr_row = 0,
        font_change_scope = 0,
    )
    функции.df_to_excel_openpyxl(
        filename = filename3,
        разновидность = "Убой ШПК",
        df_для_записи = df_бройлеры_part3,
        rowtostartin_pd = inp3,
        coltostartin_pd = 19,
        всего_colnum_offset = 2,
        неприказ_belowtablenames_offset = 0,
        приказ_belowtablenames_offset = 0,
        clearing_marker = "не удалять",
        clearing_marker_col = 1,
        clearing_offset = 1,
        remove_borders = 0,
        change_alignment = 0,
        add_borders = 0,
        aggr_row = 0,
        font_change_scope = 0,
    )
    функции.df_to_excel_openpyxl(
        filename = filename3,
        разновидность = "Убой ШПК",
        df_для_записи = df_бройлеры_part4,
        rowtostartin_pd = inp3,
        coltostartin_pd = 25,
        всего_colnum_offset = 2,
        неприказ_belowtablenames_offset = 0,
        приказ_belowtablenames_offset = 0,
        clearing_marker = "не удалять",
        clearing_marker_col = 1,
        clearing_offset = 1,
        remove_borders = 0,
        change_alignment = 0,
        add_borders = 0,
        aggr_row = 0,
        font_change_scope = 0,
    )
    end = time.time()
    print("\nушло времени на запись в накопительный отчет напрямую")
    print(end - start)
    # exit()

# СТАРКА----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
df_старка = findf.copy(deep=True)
df_старка = df_старка.drop(df_старка[(df_старка["старка"] == "нет")].index)
df_старка["направ"] = "Центр"
df_старка["комб"] = "Убой"
df_старка = функции.pd_movecol(
        df_старка,
        cols_to_move=["старка", "направ", "комб"],
        ref_col="корп",
        place="Before"
        )
df_старка = df_старка.groupby(["старка", "направ", "комб", "корп", "дата.сдачи"], as_index=False).agg({"Живок голов": "sum", "Живок вес": "sum", "Падеж голов": "sum", "Падеж вес": "sum"})
if df_реализация_fin.empty == False:
    df_старка = pd.concat([df_старка, df_реализация_fin], ignore_index = True)
#
df_старка.loc[df_старка["старка"].str.contains("Истобнянская РМ 1"), ["старка"]] = "Истобнянская (РМ 1)"
df_старка.loc[df_старка["старка"].str.contains("Истобнянская РМ 2"), ["старка"]] = "Истобнянская (РМ 2)"
df_старка.loc[df_старка["старка"].str.contains("Истобнянская РС 1"), ["старка"]] = "Истобнянская (РС 1)"
df_старка.loc[df_старка["старка"].str.contains("Истобнянская РС 2"), ["старка"]] = "Истобнянская (РС 2)"
df_старка.loc[df_старка["старка"].str.contains("Истобнянская РС 3"), ["старка"]] = "Истобнянская (РС 3)"
df_старка.loc[df_старка["старка"].str.contains("Истобнянская РС 4"), ["старка"]] = "Истобнянская (РС 4)"
df_старка.loc[df_старка["старка"].str.contains("Муромская РМ 4"), ["старка"]] = "Муромская (РМ 4)"
df_старка.loc[df_старка["старка"].str.contains("Муромская РМ 5"), ["старка"]] = "Муромская (РМ 5)"
df_старка.loc[df_старка["старка"].str.contains("Муромская РМ 7"), ["старка"]] = "Муромская (РМ 7)"
df_старка.loc[df_старка["старка"].str.contains("Муромская РС 1"), ["старка"]] = "Муромская (РС 1)"
df_старка.loc[(df_старка["старка"].str.contains("Муромская РС 6")==True) & (df_старка["старка"].str.contains(".1")==False) & (df_старка["старка"].str.contains(".2")==False), ["старка"]] = "Муромская 6 (РС 4)"
df_старка.loc[df_старка["старка"].str.contains("Муромская РС 6.1"), ["старка"]] = "Муромская 6.1 (РС 5)"
df_старка.loc[df_старка["старка"].str.contains("Муромская РС 6.2"), ["старка"]] = "Муромская 6.2 (РС 6)"
df_старка.loc[df_старка["старка"].str.contains("Муромская РС 2"), ["старка"]] = "Муромская (РС 2)"
df_старка.loc[df_старка["старка"].str.contains("Муромская РС 3"), ["старка"]] = "Муромская (РС 3)"
df_старка.loc[df_старка["старка"].str.contains("Разуменская РМ 1"), ["старка"]] = "Разуменская (РМ 1)"
df_старка.loc[df_старка["старка"].str.contains("Разуменская РМ 2"), ["старка"]] = "Разуменская (РМ 2)"
df_старка.loc[df_старка["старка"].str.contains("Разуменская РС 1"), ["старка"]] = "Разуменская (РС 1)"
df_старка.loc[df_старка["старка"].str.contains("Разуменская РС 2"), ["старка"]] = "Разуменская (РС 2)"
df_старка.loc[df_старка["старка"].str.contains("Разуменская РС 3"), ["старка"]] = "Разуменская (РС 3)"
df_старка.loc[df_старка["старка"].str.contains("Разуменская РС 4"), ["старка"]] = "Разуменская (РС 4)"
df_старка.loc[df_старка["старка"].str.contains("Тихая сосна РМ"), ["старка"]] = "Тихая сосна (РМ)"
df_старка.loc[df_старка["старка"].str.contains("Тихая сосна РС 1"), ["старка"]] = "Тихая сосна (РС 1)"
df_старка.loc[df_старка["старка"].str.contains("Тихая сосна РС 2"), ["старка"]] = "Тихая сосна (РС 2)"
#

if df_старка.empty == True:
    print("\nСТАРКИ НЕ БЫЛО")
if df_старка.empty == False:
    print("\nСТАРКА")
    print(df_старка)

if inp2 == "нет" or inp2 == "no" or inp2 == "n":
    функции.pd_toexcel(
                pd,
                #
                filename = filename0c,
                разновидность = "Лист1",
                df_для_записи = df_старка,
                header_pd = "True",
                rowtostartin_pd = 0,
                coltostartin_pd = 0,
            )
if inp2 == "да" or inp2 == "yes" or inp2 == "y":
    if df_старка.empty == False:
        df_старка_part1 = df_старка.copy(deep=True)
        df_старка_part1 = df_старка_part1.drop(["Падеж голов", "Падеж голов"], axis = 1)
        # print("\ndf_старка_part1")
        # print(df_старка_part1)

        df_старка_part2 = df_старка.copy(deep=True)
        df_старка_part2 = df_старка_part2[["Падеж голов", "Падеж голов"]]
        # print("\ndf_старка_part2")
        # print(df_старка_part2)

        функции.df_to_excel_openpyxl(
        filename = filename4,
        разновидность = "ШПК",
        df_для_записи = df_старка_part1,
        rowtostartin_pd = inp4,
        coltostartin_pd = 0,
        всего_colnum_offset = 2,
        неприказ_belowtablenames_offset = 0,
        приказ_belowtablenames_offset = 0,
        clearing_marker = "не удалять",
        clearing_marker_col = 1,
        clearing_offset = 1,
        remove_borders = 0,
        change_alignment = 0,
        add_borders = 0,
        aggr_row = 0,
        font_change_scope = 0,
    )
        функции.df_to_excel_openpyxl(
        filename = filename4,
        разновидность = "ШПК",
        df_для_записи = df_старка_part2,
        rowtostartin_pd = inp4,
        coltostartin_pd = 10,
        всего_colnum_offset = 2,
        неприказ_belowtablenames_offset = 0,
        приказ_belowtablenames_offset = 0,
        clearing_marker = "не удалять",
        clearing_marker_col = 1,
        clearing_offset = 1,
        remove_borders = 0,
        change_alignment = 0,
        add_borders = 0,
        aggr_row = 0,
        font_change_scope = 0,
    )
        