# PREPARATION
# import datetime
# import decimal
# import json
import os
# import pprint
# import re
# import shutil
import sys
# from datetime import datetime
# from decimal import Decimal
# from functools import reduce

# import numpy as np
# import openpyxl
import pandas as pd
# import sidetable
from openpyxl.styles import (Alignment, Border, Font, PatternFill, Protection,
                             Side, colors)
from openpyxl.utils import column_index_from_string, get_column_letter
from pandas.tseries.offsets import DateOffset

pd.set_option("display.max_rows", 1500)
pd.set_option("display.max_columns", 100)
pd.set_option("max_colwidth", 15)
pd.set_option("expand_frame_repr", False)
from функции import (
        json_dump_n_load,
        pd_movecol,
        # pd_readexcel,
        pd_toexcel,
        print_line,
        rawdata_budget,
        # writing_to_excel_openpyx
        )

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
# 
filename1 = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\фот рс.xlsx"
filename2a = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\фот рс - норма.xlsx"
filename2b = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\фот рс - факт.xlsx"
# 
filename3 = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\фот рм.xlsx"
filename3a = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\фот рм - норма.xlsx"
filename3b = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\фот рм - факт.xlsx"
# 
filename4 = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\рм показатели.xlsx"
filename4a = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\рм показатели свод.xlsx"
# 
filename5 = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\фот бс.xlsx"
filename5a = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\фот бс - норма.xlsx"
filename5b = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\фот бс - факт.xlsx"
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# ФОТ РС--------------------------------------------------------------------------------------------------------------
# reading from excel
df_from_excel = pd.read_excel(filename1, sheet_name="Лист1", index_col=None, engine = "openpyxl", header=0, usecols = "A,B,D,F,N,Z,AF,AG")
# print("df_from_excel")
# print(df_from_excel.head())
# должн_лист = ["Ведущий зоотехник"]
df_from_excel = df_from_excel.drop(df_from_excel[(df_from_excel["должность"] == 0)].index)
df_from_excel = df_from_excel.groupby(["тип", "площадка", "должность"], as_index=False).agg({"штат": "sum", "сц": "sum", "сумма": "sum", "поШРзн": "sum", "ия": "sum"})
df_from_excel["дес.мес"] = df_from_excel["сумма"]
# 
df_from_excel["1мес"] = df_from_excel["сумма"]/10
df_from_excel = df_from_excel.drop(["сумма"], axis = 1)
df_from_excel["на_чел_мес"] = df_from_excel["1мес"]/df_from_excel["штат"]
df_from_excel.loc[df_from_excel["штат"] < 1, ["на_чел_мес"]] = df_from_excel["на_чел_мес"]*df_from_excel["штат"]
# 
df_from_excel["поШРмес"] = df_from_excel["поШРзн"]/df_from_excel["штат"]
df_from_excel.loc[df_from_excel["штат"] < 1, ["поШРмес"]] = df_from_excel["поШРмес"]*df_from_excel["штат"]
# df_from_excel = df_from_excel.drop(["поШРзн"], axis = 1)
# 
df_from_excel["сц_на_чел"] = df_from_excel["сц"]/df_from_excel["штат"]
df_from_excel.loc[df_from_excel["штат"] < 1, ["сц_на_чел"]] = df_from_excel["сц_на_чел"]*df_from_excel["штат"]
df_from_excel["сц_на_чел_мес"] = df_from_excel["сц_на_чел"]/10
# 
# df_from_excel["всего"] = (df_from_excel["сц"]+df_from_excel["дес.мес"]+df_from_excel["ия"])/df_from_excel["штат"]
df_from_excel["всего"] = (df_from_excel["сц"]+df_from_excel["дес.мес"])/df_from_excel["штат"]
df_from_excel.loc[df_from_excel["штат"] < 1, ["всего"]] = df_from_excel["всего"]*df_from_excel["штат"]
# df_from_excel["ия"] = df_from_excel["ия"]/df_from_excel["штат"]/10
# df_from_excel.loc[df_from_excel["штат"] < 1, ["ия"]] = df_from_excel["ия"]*df_from_excel["штат"]/10
# 
df_from_excel = pd_movecol(
        df_from_excel,
        cols_to_move=["сц_на_чел"],
        ref_col="сц",
        place="After"
        )
df_from_excel = pd_movecol(
        df_from_excel,
        cols_to_move=["ия"],
        ref_col="сц_на_чел_мес",
        place="After"
        )
print("df_from_excel")
print(df_from_excel)
# sys.exit()
df_норма = df_from_excel.copy(deep=True)
df_норма = df_норма.drop(df_норма[(df_норма["тип"] == "факт")].index)
pd_toexcel(
            pd,
            # 
            df_для_записи = df_норма,
            rowtostartin_pd = 0,
            coltostartin_pd = 0,
            filename = filename2a,
            разновидность = "Лист1",
            header_pd = "True",
        )

df_факт = df_from_excel.copy(deep=True)
df_факт = df_факт.drop(df_норма[(df_норма["тип"] == "норма")].index)
pd_toexcel(
            pd,
            # 
            df_для_записи = df_факт,
            rowtostartin_pd = 0,
            coltostartin_pd = 0,
            filename = filename2b,
            разновидность = "Лист1",
            header_pd = "True",
        )

# -----------------------------------------------------------------------------------------------------------------------------------------------------
# ФОТ РМ
# reading from excel
df_from_excel = pd.read_excel(filename3, sheet_name="Лист1", index_col=None, engine = "openpyxl", header=0, usecols = "A,B,D,F,Z,AF")
# print("df_from_excel")
# print(df_from_excel)
# print(df_from_excel.dtypes)
# sys.exit()
df_from_excel = df_from_excel.drop(df_from_excel[(df_from_excel["должность"] == 0)].index)
df_from_excel = df_from_excel.groupby(["тип", "площадка", "должность"], as_index=False).agg({"штат": "sum", "сумма": "sum", "поШРзн": "sum"})
df_from_excel["1мес"] = df_from_excel["сумма"]/6
df_from_excel["на_чел_мес"] = df_from_excel["1мес"]/df_from_excel["штат"]
df_from_excel.loc[df_from_excel["штат"] < 1, ["на_чел_мес"]] = df_from_excel["на_чел_мес"]*df_from_excel["штат"]
# 
df_from_excel["поШРмес"] = df_from_excel["поШРзн"]/df_from_excel["штат"]
df_from_excel.loc[df_from_excel["штат"] < 1, ["поШРмес"]] = df_from_excel["поШРмес"]*df_from_excel["штат"]
df_from_excel = df_from_excel.drop(["поШРзн"], axis = 1)
# 
df_норма = df_from_excel.copy(deep=True)
df_норма = df_норма.drop(df_норма[(df_норма["тип"] == "факт")].index)
pd_toexcel(
            pd,
            # 
            df_для_записи = df_норма,
            rowtostartin_pd = 0,
            coltostartin_pd = 0,
            filename = filename3a,
            разновидность = "Лист1",
            header_pd = "True",
        )

df_факт = df_from_excel.copy(deep=True)
df_факт = df_факт.drop(df_норма[(df_норма["тип"] == "норма")].index)
pd_toexcel(
            pd,
            # 
            df_для_записи = df_факт,
            rowtostartin_pd = 0,
            coltostartin_pd = 0,
            filename = filename3b,
            разновидность = "Лист1",
            header_pd = "True",
        )
# РМ контрольные точки-------------------------------------------------------------------------------------------------
df_from_excel = pd.read_excel(filename4, sheet_name="Лист1", index_col=None, engine = "openpyxl", header=0, usecols = "A:J")
df_from_excel = df_from_excel.drop(df_from_excel[(df_from_excel["должность"] == 0)].index)
df_from_excel = df_from_excel.groupby(["площадка", "должность"], as_index=False).agg({
    "ЖВ 7 дней": "mean",
    "ЖВ 28 дней": "mean",
    "ЖВ 84 дня": "mean",
    "однородность 28 дней": "mean",
    "однородность 84 дня": "mean",
    "однородность по итогам тура": "mean",
    "ЖВ по итогам тура": "mean",
    "Сохранность": "mean"})
print("df_from_excel")
print(df_from_excel)
pd_toexcel(
            pd,
            # 
            df_для_записи = df_from_excel,
            rowtostartin_pd = 0,
            coltostartin_pd = 0,
            filename = filename4a,
            разновидность = "Лист1",
            header_pd = "True",
        )

# -----------------------------------------------------------------------------------------------------------------------------------------------------
# ФОТ БС

# reading from excel
df_from_excel = pd.read_excel(filename5, sheet_name="Лист1", index_col=None, engine = "openpyxl", header=0, usecols = "A,B,C,D,H,I,J,K,AF")
# print(df_from_excel.dtypes)
# sys.exit()
df_from_excel = pd_movecol(
        df_from_excel,
        cols_to_move=["поШРзн"],
        ref_col="туровая",
        place="Before"
        )
df_from_excel["1мес"] = df_from_excel["12мес"]/12
# df_from_excel.loc[df_from_excel["тип"] == "норма", ["1мес"]] = df_from_excel["3мес"]/3
# df_from_excel.loc[df_from_excel["тип"] == "факт", ["1мес"]] = df_from_excel["3мес"]/1.5
df_from_excel["на_чел_мес"] = df_from_excel["1мес"]/df_from_excel["штат"]
df_from_excel.loc[df_from_excel["штат"] < 1, ["на_чел_мес"]] = df_from_excel["на_чел_мес"]*df_from_excel["штат"]
print("df_from_excel")
print(df_from_excel)
df_норма = df_from_excel.copy(deep=True)
df_норма = df_норма.drop(df_норма[(df_норма["тип"] == "факт")].index)
pd_toexcel(
            pd,
            # 
            df_для_записи = df_норма,
            rowtostartin_pd = 0,
            coltostartin_pd = 0,
            filename = filename5a,
            разновидность = "Лист1",
            header_pd = "True",
        )
df_факт = df_from_excel.copy(deep=True)
df_факт = df_факт.drop(df_факт[(df_факт["тип"] == "норма")].index)
pd_toexcel(
            pd,
            # 
            df_для_записи = df_факт,
            rowtostartin_pd = 0,
            coltostartin_pd = 0,
            filename = filename5b,
            разновидность = "Лист1",
            header_pd = "True",
        )
