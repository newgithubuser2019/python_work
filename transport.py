# PREPARATION
import os
import datetime
from datetime import datetime
import re
import pprint
import shutil
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, colors
import json
import decimal
from decimal import Decimal
import pandas as pd
import numpy as np
import sidetable
from functools import reduce
from pandas.tseries.offsets import DateOffset
pd.set_option("display.max_rows", 1500)
pd.set_option("display.max_columns", 100)
pd.set_option("max_colwidth", 15)
pd.set_option("expand_frame_repr", False)
from функции import print_line
from функции import rawdata_budget
from функции import pd_movecol
from функции import pd_toexcel
from функции import pd_readexcel
from функции import writing_to_excel_openpyxl
from функции import json_dump_n_load
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# global variables
USERPROFILE = os.environ["USERPROFILE"]

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty lists

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty dictionaries

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty dataframes
findf = pd.DataFrame()

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
filename1 = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\база.xlsx"
filename2 = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\выгрузка.xlsx"
filename3 = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\транспорт.xlsx"
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# БАЗА--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# reading from excel
df_база = pd.read_excel(filename1, sheet_name="Лист1", index_col=None, engine = "openpyxl", header=0, usecols = "B,D")
df_база = df_база.rename(columns={"гос.номер.": "госномер"})
print("\ndf_база")
print(df_база.head())
print(df_база.shape)
# print(df_база)
print("------------------------------------------------------")

# ВЫГРУЗКА--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# reading from excel
df_выгрузка = pd.read_excel(filename2, sheet_name="Лист1", index_col=None, engine = "openpyxl", header=0, usecols = "A,B,C,E,F,I,L")
df_выгрузка = df_выгрузка.rename(columns={"Гос. номер": "госномер"})
df_выгрузка["госномер"]=df_выгрузка["госномер"].astype(str)
df_выгрузка["Год выпуска"] = pd.to_numeric(df_выгрузка["Год выпуска"], errors="coerce")
df_выгрузка = df_выгрузка.fillna("X")
df_выгрузка["госномер"] = df_выгрузка["госномер"].str.replace(" ","")
df_выгрузка["госномер"] = df_выгрузка["госномер"].str.replace("RUSБ","")
df_выгрузка["госномер"] = df_выгрузка["госномер"].str.replace("RUS","")
# print(df_выгрузка)
# exit()
df_выгрузка["тип"] = df_выгрузка["Тип ТС"]
df_выгрузка["перевозит"] = ""
# 
df_выгрузка.loc[df_выгрузка["тип"].str.contains("трактор", na=False), ["тип"]] = "трактор"
df_выгрузка.loc[df_выгрузка["тип"].str.contains("Легк", na=False, case=False), ["тип"]] = "легковой"
df_выгрузка.loc[df_выгрузка["тип"].str.contains("цистер", na=False, case=False), ["тип"]] = "цистерна"
df_выгрузка.loc[df_выгрузка["тип"].str.contains("самосв", na=False, case=False), ["тип"]] = "самосвал"
df_выгрузка.loc[df_выгрузка["тип"].str.contains("проч", na=False, case=False), ["тип"]] = "прочие"
df_выгрузка.loc[df_выгрузка["тип"].str.contains("спецтех", na=False, case=False), ["тип"]] = "проч.спец.техника"
df_выгрузка.loc[df_выгрузка["тип"].str.contains("автобус", na=False, case=False), ["тип"]] = "автобус"
df_выгрузка.loc[df_выгрузка["тип"].str.contains("легко|автобус", na=False, case=False), ["перевозит"]] = "персонал"
df_выгрузка.loc[df_выгрузка["тип"].str.contains("борт", na=False, case=False), ["тип"]] = "бортовой"
df_выгрузка.loc[df_выгрузка["тип"].str.contains("тягач", na=False, case=False), ["тип"]] = "проч.спец.техника"
df_выгрузка.loc[df_выгрузка["тип"].str.contains("погруз", na=False, case=False), ["тип"]] = "погрузчик"
df_выгрузка.loc[df_выгрузка["Подразделение"].str.contains("персонала", na=False), ["перевозит"]] = "персонал"
df_выгрузка.loc[df_выгрузка["Подразделение"].str.contains("БКЗ|ШКЗ|Корм", na=False, case=False), ["перевозит"]] = "кормовоз"
df_выгрузка.loc[df_выгрузка["Подразделение"].str.contains("птиц", na=False, case=False), ["тип"]] = "птицевоз"
df_выгрузка.loc[df_выгрузка["Подразделение"].str.contains("цыпл", na=False, case=False), ["тип"]] = "яйцевоз+цып"
df_выгрузка.loc[df_выгрузка["Подразделение"].str.contains("инкуб", na=False, case=False), ["тип"]] = "яйцевоз+цып"
df_выгрузка.loc[df_выгрузка["Наименование"].str.contains("легк|автобус", na=False, case=False), ["перевозит"]] = "персонал"
df_выгрузка.loc[df_выгрузка["Наименование"].str.contains("трактор", na=False, case=False), ["тип"]] = "трактор"
df_выгрузка.loc[df_выгрузка["Наименование"].str.contains("борт", na=False, case=False), ["тип"]] = "бортовой"
df_выгрузка.loc[df_выгрузка["Наименование"].str.contains("дезинф", na=False, case=False), ["тип"]] = "проч.спец.техника"
df_выгрузка.loc[df_выгрузка["Наименование"].str.contains("автобус", na=False, case=False), ["тип"]] = "автобус"
df_выгрузка.loc[df_выгрузка["Наименование"].str.contains("самосв", na=False, case=False), ["тип"]] = "самосвал"
df_выгрузка.loc[df_выгрузка["Наименование"].str.contains("дук", na=False, case=False), ["тип"]] = "проч.спец.техника"
df_выгрузка.loc[df_выгрузка["Наименование"].str.contains("легк", na=False, case=False), ["тип"]] = "легковой"
df_выгрузка.loc[df_выгрузка["Наименование"].str.contains("экск|троту|шасси|вакуум", na=False, case=False), ["тип"]] = "проч.спец.техника"
df_выгрузка.loc[df_выгрузка["Наименование"].str.contains("дук", na=False, case=False), ["перевозит"]] = "прочие тех.и хоз.работы"
df_выгрузка.loc[df_выгрузка["Наименование"].str.contains("погруз", na=False, case=False), ["тип"]] = "погрузчик"
df_выгрузка.loc[(df_выгрузка["Наименование"].str.contains("грузов", na=False, case=False)) & (df_выгрузка["тип"].str.contains("X", na=False, case=False)), ["тип"]] = "птицевоз"
df_выгрузка.loc[(df_выгрузка["Наименование"].str.contains("грузов", na=False, case=False)) & (df_выгрузка["тип"].str.contains("X", na=False, case=False)), ["перевозит"]] = "перевозка ЖВ"
df_выгрузка.loc[df_выгрузка["перевозит"].str.contains("корм", na=False, case=False), ["тип"]] = "кормовоз"
# 
df_выгрузка.loc[df_выгрузка["тип"].str.contains("бортовой|птицевоз|трактор|погруз", na=False, case=False), ["перевозит"]] = "перевозка ЖВ"
df_выгрузка.loc[df_выгрузка["тип"].str.contains("яйцевоз", na=False, case=False), ["перевозит"]] = "ия+сц"
df_выгрузка.loc[df_выгрузка["тип"].str.contains("самосвал|спец|цист", na=False, case=False), ["перевозит"]] = "прочие тех.и хоз.работы"
# госномер_из_базы = df_база["госномер"].tolist()
# for i in госномер_из_базы:
    # df_выгрузка.loc[df_выгрузка["госномер"].str.contains(str(i), na=False), ["госномер"]] = i
print("\ndf_выгрузка")
print(df_выгрузка.head())
print(df_выгрузка.shape)
# print(df_выгрузка)
print("------------------------------------------------------")

# merging dataframes
df_from_excel = pd.merge(df_база, df_выгрузка, how = "left", on = ["госномер"])
print("df_from_excel")
print(df_from_excel)
pd_toexcel(
            pd,
            # 
            df_для_записи = df_from_excel,
            rowtostartin_pd = 0,
            coltostartin_pd = 0,
            filename = filename3,
            разновидность = "Лист1",
            header_pd = "True",
        )
