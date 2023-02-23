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
filename0 = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\стаж на 14.09.2021.xlsx"
filename1 = USERPROFILE + "\\Documents\\Работа\\отдельные поручения\\стаж группировка.xlsx"
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# стаж
# reading from excel
df_from_excel = pd.read_excel(filename0, sheet_name="Лист1", index_col=0, engine = "openpyxl", header=0) # pd_read_excel_cols_list)
df_from_excel.reset_index(inplace = True)
df_from_excel = df_from_excel.drop(["№ п/п"], axis = 1)
df_from_excel = df_from_excel.fillna(0)
df_from_excel["до1"] = ""
df_from_excel["от1до2"] = ""
df_from_excel["больше2"] = ""
df_from_excel.loc[df_from_excel["Стаж"] <= 1, ["до1"]] = "X"
df_from_excel.loc[((df_from_excel["Стаж"] > 1) & (df_from_excel["Стаж"] <= 2)), ["от1до2"]] = "X"
df_from_excel.loc[df_from_excel["Стаж"] > 2, ["больше2"]] = "X"
print("df_from_excel")
print(df_from_excel)
# for col in df_from_excel.columns:
    # print(col)

# sidetable section
df_sidetable1 = df_from_excel.stb.freq(["Подразделение", "до1"])
df_sidetable1 = df_sidetable1[["Подразделение", "до1", "count"]]
df_sidetable1.drop(df_sidetable1.loc[df_sidetable1["до1"] == ""].index, inplace=True)
df_sidetable1 = df_sidetable1.rename(columns={df_sidetable1.columns[2]: "кол-во_до1"})
df_sidetable2 = df_from_excel.stb.freq(["Подразделение", "от1до2"])
df_sidetable2 = df_sidetable2[["Подразделение", "от1до2", "count"]]
df_sidetable2.drop(df_sidetable2.loc[df_sidetable2["от1до2"] == ""].index, inplace=True)
df_sidetable2 = df_sidetable2.rename(columns={df_sidetable2.columns[2]: "кол-во_от1до2"})
df_sidetable3 = df_from_excel.stb.freq(["Подразделение", "больше2"])
df_sidetable3 = df_sidetable3[["Подразделение", "больше2", "count"]]
df_sidetable3.drop(df_sidetable3.loc[df_sidetable3["больше2"] == ""].index, inplace=True)
df_sidetable3 = df_sidetable3.rename(columns={df_sidetable3.columns[2]: "кол-во_больше2"})
DFs_to_merge = [df_sidetable1, df_sidetable2, df_sidetable3]
df_sidetable = reduce(lambda left, right: pd.merge(left, right, on = "Подразделение", how="outer"), DFs_to_merge)
df_sidetable = df_sidetable.drop(["до1"], axis = 1)
df_sidetable = df_sidetable.drop(["от1до2"], axis = 1)
df_sidetable = df_sidetable.drop(["больше2"], axis = 1)
print("\ndf_sidetable")
print(df_sidetable)

# writing to excel
pd_toexcel(
            pd,
            # 
            df_для_записи = df_sidetable,
            rowtostartin_pd = 0,
            coltostartin_pd = 0,
            filename = filename1,
            разновидность = "Лист1",
            header_pd = "True",
        )
# exit()
