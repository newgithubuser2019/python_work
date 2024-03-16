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
import plotly.express as px

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
# filename5 = "D:\\programming\\_datasets\\просмотры резюме.xlsx"
filename0 = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\Тест excel.xlsm"
filename0b = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\задание 6 - df.xlsx"
filename2a = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\задание 2 - федеральный округ.xlsx"
filename2b = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\задание 2 - функциональный блок.xlsx"

# ------------------------------------------------------------------------------------------------------------------------------------------
df_from_excel = pd.read_excel(
    filename0,
    sheet_name="Штатное расписание",
    # index_col=0,
    # engine = "openpyxl",
    header=2,
    # usecols = "A:L",
    usecols = "A,C,G,H,I,K",
    )
print(df_from_excel)
# sys.exit()
