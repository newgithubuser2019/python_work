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
# filename0a = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\_промежуточный файл df_впк.xlsx"
# filename = USERPROFILE + "\\Documents\\Google Sheets Test.xlsx"
# filename5 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\время поднятия кормушки\\время поднятия кормушки.xlsx"
# filename5 = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\время поднятия кормушки.xlsx"
filename5 = "D:\\programming\\_datasets\\просмотры резюме.xlsx"

# ------------------------------------------------------------------------------------------------------------------------------------------
df_from_excel = pd.read_excel(
    filename5,
    sheet_name="Лист1",
    # index_col=0,
    # engine = "openpyxl",
    header=0,
    # usecols = "A,B",
    )
# df_from_excel = df_from_excel.loc[(df_from_excel["Вид выбытия"].str.contains("Основная")) | (df_from_excel["Вид выбытия"].str.contains("Разрежение"))]
# df_from_excel = df_from_excel.dropna(subset=["Время поднятия кормушки"])
df_from_excel["компания"] = "пусто"
df_from_excel.loc[df_from_excel["initial"].str.contains(","), ["компания"]] = df_from_excel["initial"]
df_from_excel["дата"] = None
df_from_excel.loc[df_from_excel["компания"] == "пусто", ["дата"]] = df_from_excel["initial"]
df_from_excel = df_from_excel.drop(["initial"], axis = 1)
df_from_excel["дата"] = df_from_excel["дата"].ffill()
df_from_excel = df_from_excel.drop(df_from_excel.loc[df_from_excel["компания"]=="пусто"].index)
df_from_excel.loc[df_from_excel["компания"].str.contains(","), ["компания"]] = df_from_excel["компания"].str.rsplit(",").str[0]
df_from_excel = df_from_excel.groupby(["компания"], as_index=False).count()
df_from_excel = df_from_excel.sort_values(by=["дата"], ascending=False)
print(df_from_excel)
sys.exit()

