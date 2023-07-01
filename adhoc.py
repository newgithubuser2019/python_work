# IMPORTS
import datetime
import decimal
import json
import os
import pprint
import re
import shutil
from functools import reduce

import numpy as np
import openpyxl
import pandas as pd
import sidetable
from pandas.tseries.offsets import DateOffset

import функции

pd.set_option("display.max_rows", 1500)
pd.set_option("display.max_columns", 100)
pd.set_option("max_colwidth", 15)
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
# path1 = USERPROFILE + "\\Documents\\Работа\\отчетность\\выход гп\\акты\\"
# listoffiles = os.listdir(path1)
# filename0 = USERPROFILE + "\\Documents\\Работа\\отчетность\\накопительный отчет\\Накопительный (старка) Белгород 2022.xlsx"
filename0 = USERPROFILE + "\\Documents\\Google Sheets Test.xlsx"

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
SHEET_ID = '1Y56CvHeHegMYlx1V-bzC9jc7cqRf7tLgRcIvRKbn5sk'
SHEET_NAME = 'Subscriptions'
url = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}'
#df = pd.read_excel(url, header=0, usecols="A,B,C,D", engine="openpyxl")
df = pd.read_csv(url, header=0)
print(df.head())
df = df["Full Name", "Amount"]
print(df.head())
"""
df = pd.read_excel(
                filename0,
                sheet_name="Subscriptions",
                index_col=None,
                engine = "openpyxl",
                header=0,
                usecols = "A:G",
                #dtype = {"I": str},
                )
df["First Name"] = df["Full Name"].str.rsplit(" ").str[0]
df["Last Name"] = df["Full Name"].str.rsplit(" ").str[-1]
print(df.head())
#df_active = df.drop(df.loc[df["Status"]=="Inactive"].index)
df_active = df.loc[df["Status"]=="Active"]
print(df_active)
