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
count = 0
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# default lists

# empty dataframes
findf = pd.DataFrame()
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# file paths
# for корпус in ["1-6", "7-11"]:
listoffiles = os.listdir(USERPROFILE + "\\Documents\\Работа\\закрытие зп\\за тур\\выращивание\\Муромская\\4\\")
filename0a = USERPROFILE + "\\Documents\\Работа\\закрытие зп\\за тур\\выращивание\\Муромская\\свод 4.xlsx"
# filename0b = USERPROFILE + "\\Documents\\Работа\\оу контрольная карта убой\\_промежуточный файл findf_1.xlsx"
# filename0c = USERPROFILE + "\\Documents\\Работа\\оу контрольная карта убой\\_промежуточный файл findf_2.xlsx"
# filename1 = USERPROFILE + "\\Documents\\Работа\\оу контрольная карта убой\\Динамика развития птицы 2021 Белгород.xlsx"
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# exit()
for i in listoffiles:
    count += 1
    print(str(count))
    месяц = i[5:-5]
    print(месяц)
    """
    regexpr = re.compile(r"(\d{1,4})+(.)")
    датасдачи = ""
    for gr in regexpr.findall(i):
        датасдачи = датасдачи + gr[0]
        датасдачи = датасдачи + gr[1]
    датасдачи = датасдачи[0:-1]
    print(датасдачи)
    today = datetime.strptime(датасдачи, "%d.%m.%Y")
    weeknum = datetime.date(today).isocalendar().week
    """
    # exit()
    # кку---------------------------------------------------------------------
    df_from_excel = pd.read_excel(USERPROFILE + "\\Documents\\Работа\\закрытие зп\\за тур\\выращивание\\Муромская\\4\\" + i, sheet_name="Лист1", index_col=0, engine = "openpyxl", header=0, usecols = "A,B,G") # pd_read_excel_cols_list)
    df_from_excel.reset_index(inplace = True)
    df_from_excel = df_from_excel.drop(df_from_excel[(df_from_excel["должность"] == "Санитар ветеринарный")].index)
    df_from_excel = df_from_excel.drop(["должность"], axis = 1)
    df_from_excel = df_from_excel.rename(columns={"явки": месяц})
    # df_from_excel["месяц"] = месяц
    print(df_from_excel)
    # df_свод = df_from_excel.copy(deep=True)
    # df_свод = pd.merge(df_свод, df_from_excel, how = "left", on = ["ФИО2", "должность"])
    # print(df_свод)
    if count == 1:
        df_from_excel2 = df_from_excel.copy(deep=True)
        df_from_excel2 = df_from_excel2.drop(df_from_excel2.columns[[1]], axis = 1)
        findf = findf.append(df_from_excel2, ignore_index = True)
    findf = pd.merge(findf, df_from_excel, how = "outer", on = ["ФИО2"])
    # exit()
    # print(findf)
# df_pivot = pd.pivot_table(findf, index=["месяц", "ФИО2", "должность"], values=["явки"])
# print(df_pivot)
pd_toexcel(
            pd,
            # 
            df_для_записи = findf,
            rowtostartin_pd = 0,
            coltostartin_pd = 0,
            filename = filename0a,
            разновидность = "Лист1",
            header_pd = "True",
        )