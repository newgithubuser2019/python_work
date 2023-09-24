# IMPORTS
import datetime
import decimal
import json
import os
import pprint
import re
import shutil
import sys
from functools import reduce

import numpy as np
import openpyxl
import pandas as pd
# import sidetable
from pandas.tseries.offsets import DateOffset

import функции

pd.set_option("display.max_rows", 1500)
pd.set_option("display.max_columns", 100)
pd.set_option("max_colwidth", 30)
pd.set_option("expand_frame_repr", False)
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# global variables
USERPROFILE = os.environ["USERPROFILE"]

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty lists

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty dictionaries

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty dataframes
# findf = pd.DataFrame()
df_bez_korma = pd.DataFrame()

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# default lists

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# default dictionaries

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# prompts for user input
# prompt1 = "\nРеализация: "

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# user inputs
# inp1 = input(prompt1)

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# file paths
# path_1 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\время поднятия кормушки\\"
# listoffiles_1 = os.listdir(path_1)
# file names
filename0a = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\_промежуточный файл df_впк.xlsx"
# filename = USERPROFILE + "\\Documents\\Google Sheets Test.xlsx"
# filename5 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\время поднятия кормушки\\время поднятия кормушки.xlsx"
filename5 = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\время поднятия кормушки.xlsx"


# ------------------------------------------------------------------------------------------------------------------------------------------
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
# sys.exit()
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
print("\ndf_pivot")
print(df_pivot)
# sys.exit()
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
