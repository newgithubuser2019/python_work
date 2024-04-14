# IMPORTS
import os

# import sidetable
import sys

import rich
from rich.console import Console
from rich.traceback import install

install(suppress=[rich], show_locals=False)
console = Console()
# import numpy as np
import pandas as pd

import функции

pd.set_option("display.max_rows", 1500)
pd.set_option("display.max_columns", 100)
pd.set_option("max_colwidth", 15)
pd.set_option("expand_frame_repr", False)
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# global variables
USERPROFILE = os.environ["USERPROFILE"]

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty lists

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty dictionaries

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty dataframes
findf = pd.DataFrame()

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# default lists

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# default dictionaries

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# prompts for user input
prompt0 = "\nцб или старка?: "
# prompt1 = "\nдетализация до корпусов?: "

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# user inputs
inp0 = input(prompt0)
# inp1 = input(prompt1)
inp1 = "нет"

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# file paths
filename1 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\Накопительный 2024 - Белгород.xlsx"
filename2 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\оу поступление живка\\цб.xlsx"
filename3 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\Накопительный (старка) Белгород 2024.xlsx"
filename4 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\оу поступление живка\\старка.xlsx"

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# накопительный отчет - сравнение
дата_нач = "01.01.2024"
дата_нач = pd.to_datetime(дата_нач, dayfirst=True)
дата_кон = "31.12.2024"
дата_кон = pd.to_datetime(дата_кон, dayfirst=True)

# ЦБ-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
if inp0 == "цб":
        # загрузка накопительного отчета-----------------------------------------------------------------------------------------------------------------------------------------------------
        df_цб = pd.read_excel(
                filename1,
                sheet_name="Убой ШПК",
                index_col=0,
                engine = "openpyxl",
                header=6,
                # usecols = "B,I,O,V,W",
                usecols = "B,I,O,V,W,Z,AA",
                dtype = {"I": str},
                )
        df_цб.reset_index(inplace = True)
        df_цб = df_цб.drop([0], axis = 0) # drop row with index 0
        df_цб = df_цб.rename(columns={
                "Площадка": "ОП",
                " корпуса": "корп",
                "Дата сдачи": "д_сдачи",
                "Кол-во голов без падежа": "гол",
                "Общий вес без падежа": "вес",
                "выбраковка с/н, гол.": "выбр_г",
                "выбраковка с/н, кг": "выбр_в",
                })
        df_цб = df_цб.dropna(subset=["ОП"])
        df_цб["д_сдачи"] = pd.to_datetime(df_цб["д_сдачи"], dayfirst=True)
        df_цб = df_цб.drop(df_цб[(df_цб["д_сдачи"] < дата_нач)].index)
        df_цб = df_цб.drop(df_цб[(df_цб["д_сдачи"] > дата_кон)].index)
        df_цб["выбр_г"] = df_цб["выбр_г"].fillna(0)
        df_цб["выбр_в"] = df_цб["выбр_в"].fillna(0)
        # print("\ndf_цб")
        # print(df_цб)
        # sys.exit()
        df_цб["гол"] = df_цб["гол"] - df_цб["выбр_г"]*0 # вычитается в НО
        df_цб["вес"] = df_цб["вес"] - df_цб["выбр_в"]*0 # # вычитается в НО
        df_цб = df_цб.drop(["выбр_г", "выбр_в"], axis = 1)
        if inp1 == "да" or inp1 == "yes" or inp1 == "y":
                df_цб = df_цб.groupby(["д_сдачи", "ОП", "корп"], as_index=False).agg({"гол": "sum", "вес": "sum"})
                df_цб.loc[df_цб["ОП"].str.contains("Муром"), ["корп"]] = df_цб["корп"]*10
        if inp1 == "нет" or inp1 == "no" or inp1 == "n":
                df_цб = df_цб.groupby(["д_сдачи", "ОП"], as_index=False).agg({"гол": "sum", "вес": "sum"})
        try:
                df_цб = df_цб.sort_values(by=["д_сдачи", "ОП", "корп"], ascending=True)
        except KeyError:
                df_цб = df_цб.sort_values(by=["д_сдачи", "ОП"], ascending=True)
        df_цб.reset_index(inplace = True)
        df_цб = df_цб.drop(["index"], axis = 1)
        print("\ndf_цб")
        print(df_цб)
        гол_сумма = df_цб["гол"].sum()
        print(гол_сумма)
        вес_сумма = df_цб["вес"].sum()
        print(вес_сумма)
        # sys.exit()

        # загрузка оу поступление живка-----------------------------------------------------------------------------------------------------------------------------------------------------
        df_поступление = pd.read_excel(
                filename2,
                sheet_name="TDSheet",
                index_col=0,
                engine = "openpyxl",
                header=9,
                usecols = "A,D,F,S,T",
                # usecols = "A,D,G,S,T",
                dtype = {"T": object},
                )
        df_поступление.reset_index(inplace = True)
        df_поступление = df_поступление.rename(columns={"Дата": "д_сдачи", "Поставщик (Площадка)": "ОП", "Корпус": "корп", "Поступило по ТТН Голов, гол.": "ттн_гол", "Поступило по ТТН Вес, кг.": "ттн_вес"})
        df_поступление = df_поступление.dropna(subset=["ОП"])
        df_поступление = функции.pd_movecol(
                df_поступление,
                cols_to_move=["д_сдачи"],
                ref_col="корп",
                place="After"
                )
        df_поступление["д_сдачи"] = pd.to_datetime(df_поступление["д_сдачи"], dayfirst=True)
        df_поступление = df_поступление.drop(df_поступление[(df_поступление["д_сдачи"] < дата_нач)].index)
        df_поступление = df_поступление.drop(df_поступление[(df_поступление["д_сдачи"] > дата_кон)].index)
        df_поступление.loc[df_поступление["ОП"].str.contains("Агрин"), ["ОП"]] = "Агрин"
        df_поступление.loc[df_поступление["ОП"].str.contains("Коренская"), ["ОП"]] = "Коренское"
        df_поступление.loc[df_поступление["ОП"].str.contains("Графовская"), ["ОП"]] = "Графовское"
        df_поступление.loc[df_поступление["ОП"].str.contains("Полянская"), ["ОП"]] = "Полянское"
        df_поступление.loc[df_поступление["ОП"].str.contains("Томаровская"), ["ОП"]] = "Томаровское"
        # df_поступление.loc[(df_поступление["ОП"].str.contains("Муромская")==True) & (df_поступление["ОП"].str.contains(" РМ")==False), ["ОП"]] = "Муромское"
        df_поступление.loc[(df_поступление["ОП"].str.contains("Муромская 1")==True) & (df_поступление["ОП"].str.contains(" РМ")==False), ["ОП"]] = "Муромское 1"
        df_поступление.loc[(df_поступление["ОП"].str.contains("Муромская 2")==True) & (df_поступление["ОП"].str.contains(" РМ")==False), ["ОП"]] = "Муромское 2"
        df_поступление.loc[df_поступление["ОП"].str.contains("Нежегольская"), ["ОП"]] = "Нежегольское"
        df_поступление.loc[df_поступление["ОП"].str.contains("Валуйская"), ["ОП"]] = "Валуйское"
        df_поступление.loc[df_поступление["ОП"].str.contains("Рождественская"), ["ОП"]] = "Рождественское"
        df_поступление.loc[df_поступление["корп"].str.contains("корпус"), ["корп"]] = df_поступление["корп"].str[0:3]
        df_поступление.loc[df_поступление["корп"].str.contains(" к", na=False), ["корп"]] = df_поступление["корп"].str[0:2]
        if inp1 == "да" or inp1 == "yes" or inp1 == "y":
                df_поступление = df_поступление.groupby(["д_сдачи", "ОП", "корп"], as_index=False).agg({"ттн_гол": "sum", "ттн_вес": "sum"})
                df_поступление["корп"] = pd.to_numeric(df_поступление["корп"], errors="coerce")
        if inp1 == "нет" or inp1 == "no" or inp1 == "n":
                df_поступление = df_поступление.groupby(["д_сдачи", "ОП"], as_index=False).agg({"ттн_гол": "sum", "ттн_вес": "sum"})
        try:
                df_поступление = df_поступление.sort_values(by=["д_сдачи", "ОП", "корп"], ascending=True)
        except KeyError:
                df_поступление = df_поступление.sort_values(by=["д_сдачи", "ОП"], ascending=True)
        df_поступление.reset_index(inplace = True)
        df_поступление = df_поступление.drop(["index"], axis = 1)
        print("\ndf_поступление")
        print(df_поступление)
        гол_сумма = df_поступление["ттн_гол"].sum()
        print(гол_сумма)
        вес_сумма = df_поступление["ттн_вес"].sum()
        print(вес_сумма)
        # sys.exit()

        # merging dataframes---------------------------------------------------------------------------------------------------------------------------------------------------
        try:
                df_цб = pd.merge(df_цб, df_поступление, how = "outer", on = ["д_сдачи", "ОП", "корп"])
                df_цб = df_цб.sort_values(by=["д_сдачи", "ОП", "корп"], ascending=True)
        except KeyError:
                df_цб = pd.merge(df_цб, df_поступление, how = "outer", on = ["д_сдачи", "ОП"])
                df_цб = df_цб.sort_values(by=["д_сдачи", "ОП"], ascending=True)
        # df_цб = df_цб.sort_values(by=["д_сдачи", "ОП"], ascending=True)
        df_цб["НО-1С_гол"] = df_цб["гол"] - df_цб["ттн_гол"]
        df_цб["НО-1С_вес"] = df_цб["вес"] - df_цб["ттн_вес"]
        df_цб = df_цб.drop(df_цб[(df_цб["НО-1С_гол"] == 0) & (df_цб["НО-1С_вес"] == 0)].index)
        # df_цб.loc[df_цб["НО-1С_вес"] < 0, ["НО-1С_вес"]] = df_цб["НО-1С_вес"]*(-1) # зачем я умножаю?
        df_цб = df_цб.drop(df_цб[(df_цб["НО-1С_гол"] == 0) & (df_цб["НО-1С_вес"] < 0.000000001)].index)
        функции.print_line("hyphens")
        print("\nСРАВНЕНИЕ")
        if df_цб.empty == True:
                print("ВСЕ СХОДИТСЯ")
        if df_цб.empty == False:
                print(df_цб)

# СТАРКА-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
if inp0 == "старка":
        # загрузка накопительного отчета---------------------------------------------------------------------------------------------------------------------------------------------------
        df_starka = pd.read_excel(
                filename3,
                sheet_name="ШПК",
                index_col=0,
                engine = "openpyxl",
                header=7,
                usecols = "B,D,E,F,G,H,L,M",
                dtype = {"E": str},
                )
        df_starka.reset_index(inplace = True)
        df_starka = df_starka.rename(columns={1: "ОП", 2: "сдача", "2.1": "корп", 3: "д_сдачи", 4: "гол", 5: "вес", 9: "п_гол", 10: "п_вес"})
        df_starka = df_starka.dropna(subset=["ОП"])
        df_starka["гол"] = df_starka["гол"] - df_starka["п_гол"]
        df_starka["вес"] = df_starka["вес"] - df_starka["п_вес"]
        df_starka = df_starka.drop(["п_гол"], axis = 1)
        df_starka = df_starka.drop(["п_вес"], axis = 1)
        # df_starka = df_starka.drop(df_starka[(df_starka["сдача"] == "Реализация")].index)
        df_starka = df_starka.loc[df_starka["сдача"].str.contains("Убой")]
        df_starka = df_starka.drop(["сдача"], axis = 1)
        df_starka.loc[df_starka["ОП"].str.contains("Истобнянская"), ["ОП"]] = "Истобнянская"
        df_starka.loc[df_starka["ОП"].str.contains("Муромская"), ["ОП"]] = "Муромская"
        df_starka.loc[df_starka["ОП"].str.contains("Разуменская"), ["ОП"]] = "Разуменская"
        df_starka.loc[df_starka["ОП"].str.contains("Тихая сосна"), ["ОП"]] = "Тихая сосна"
        df_starka["д_сдачи"] = pd.to_datetime(df_starka["д_сдачи"], dayfirst=True)
        df_starka = df_starka.drop(df_starka[(df_starka["д_сдачи"] < дата_нач)].index)
        df_starka = df_starka.drop(df_starka[(df_starka["д_сдачи"] > дата_кон)].index)
        if inp1 == "да" or inp1 == "yes" or inp1 == "y":
                df_starka = df_starka.groupby(["д_сдачи", "ОП", "корп"], as_index=False).agg({"гол": "sum", "вес": "sum"})
        if inp1 == "нет" or inp1 == "no" or inp1 == "n":
                df_starka = df_starka.groupby(["д_сдачи", "ОП"], as_index=False).agg({"гол": "sum", "вес": "sum"})
        try:
                df_starka = df_starka.sort_values(by=["д_сдачи", "ОП", "корп"], ascending=True)
                #
                df_starka["корп"] = df_starka["корп"].astype(str)+"_"
                df_starka["корп"] = "_"+df_starka["корп"]
                # df_starka["корп"] = df_starka["корп"].map(lambda x: x.rstrip(" ")) # здесь не пробел, а специальный символ из 1С
                # df_starka["корп"] = df_starka["корп"].str.replace(" ","") # здесь не пробел, а специальный символ из 1С
                df_starka["корп"] = df_starka["корп"].apply(lambda x: x.replace(" ","")) # здесь не пробел, а специальный символ из 1С
                df_starka["корп"] = df_starka["корп"].apply(lambda x: x.replace("_",""))
                # df_starka["корп"] = df_starka["корп"].apply(lambda x: float(x) if str(x).isdigit() else x)
                # df_starka["корп2"] = df_starka["корп"].dtype
        except KeyError:
                df_starka = df_starka.sort_values(by=["д_сдачи", "ОП"], ascending=True)
        df_starka.reset_index(inplace = True)
        df_starka = df_starka.drop(["index"], axis = 1)
        #
        print("\ndf_starka")
        print(df_starka)
        # print(df_starka.head())
        # print(df_starka.dtypes)
        гол_сумма = df_starka["гол"].sum()
        print(гол_сумма)
        вес_сумма = df_starka["вес"].sum()
        print(вес_сумма)
        # sys.exit()

        # загрузка оу поступление живка-----------------------------------------------------------------------------------------------------------------------------------------------------
        df_поступление_starka = pd.read_excel(
                filename4,
                sheet_name="TDSheet",
                index_col=0,
                engine = "openpyxl",
                header=9,
                usecols = "A,D,F,S,T",
                dtype = {"T": object, "F": str},
                )
        df_поступление_starka.reset_index(inplace = True)
        df_поступление_starka = df_поступление_starka.rename(columns={"Дата": "д_сдачи", "Поставщик (Площадка)": "ОП", "Корпус": "корп", "Поступило по ТТН Голов, гол.": "ттн_гол", "Поступило по ТТН Вес, кг.": "ттн_вес"})
        df_поступление_starka = df_поступление_starka.dropna(subset=["ОП"])
        df_поступление_starka = функции.pd_movecol(
                df_поступление_starka,
                cols_to_move=["д_сдачи"],
                ref_col="корп",
                place="After"
                )
        df_поступление_starka["д_сдачи"] = pd.to_datetime(df_поступление_starka["д_сдачи"], dayfirst=True)
        df_поступление_starka = df_поступление_starka.drop(df_поступление_starka[(df_поступление_starka["д_сдачи"] < дата_нач)].index)
        df_поступление_starka = df_поступление_starka.drop(df_поступление_starka[(df_поступление_starka["д_сдачи"] > дата_кон)].index)
        df_поступление_starka.loc[df_поступление_starka["ОП"].str.contains("Истобнянская"), ["ОП"]] = "Истобнянская"
        df_поступление_starka.loc[df_поступление_starka["ОП"].str.contains("Муромская"), ["ОП"]] = "Муромская"
        df_поступление_starka.loc[df_поступление_starka["ОП"].str.contains("Разуменская"), ["ОП"]] = "Разуменская"
        df_поступление_starka.loc[df_поступление_starka["ОП"].str.contains("Тихая сосна"), ["ОП"]] = "Тихая сосна"
        df_поступление_starka.loc[df_поступление_starka["корп"].str.contains("корпус"), ["корп"]] = df_поступление_starka["корп"].str[0:3]
        df_поступление_starka.loc[df_поступление_starka["корп"].str.contains(" к", na=False), ["корп"]] = df_поступление_starka["корп"].str[0:2]
        if inp1 == "да" or inp1 == "yes" or inp1 == "y":
                df_поступление_starka = df_поступление_starka.groupby(["д_сдачи", "ОП", "корп"], as_index=False).agg({"ттн_гол": "sum", "ттн_вес": "sum"})
        if inp1 == "нет" or inp1 == "no" or inp1 == "n":
                df_поступление_starka = df_поступление_starka.groupby(["д_сдачи", "ОП"], as_index=False).agg({"ттн_гол": "sum", "ттн_вес": "sum"})
        try:
                df_поступление_starka = df_поступление_starka.sort_values(by=["д_сдачи", "ОП", "корп"], ascending=True)
                #
                df_поступление_starka["корп"] = df_поступление_starka["корп"].astype(str)+"_"
                df_поступление_starka["корп"] = "_"+df_поступление_starka["корп"]
                # df_поступление_starka["корп"] = df_поступление_starka["корп"].map(lambda x: x.rstrip(" ")) # здесь не пробел, а специальный символ из 1С
                # df_поступление_starka["корп"] = df_поступление_starka["корп"].str.replace(" ","") # здесь не пробел, а специальный символ из 1С
                df_поступление_starka["корп"] = df_поступление_starka["корп"].apply(lambda x: x.replace(" ","")) # здесь не пробел, а специальный символ из 1С
                df_поступление_starka["корп"] = df_поступление_starka["корп"].apply(lambda x: x.replace("_",""))
                # df_поступление_starka["корп"] = df_поступление_starka["корп"].apply(lambda x: float(x) if str(x).isdigit() else x)
                # df_поступление_starka["корп2"] = df_поступление_starka["корп"].dtype
        except KeyError:
                df_поступление_starka = df_поступление_starka.sort_values(by=["д_сдачи", "ОП"], ascending=True)
        df_поступление_starka.reset_index(inplace = True)
        df_поступление_starka = df_поступление_starka.drop(["index"], axis = 1)
        #
        print("\ndf_поступление_starka")
        print(df_поступление_starka)
        # print(df_поступление_starka.dtypes)
        гол_сумма = df_поступление_starka["ттн_гол"].sum()
        print(гол_сумма)
        вес_сумма = df_поступление_starka["ттн_вес"].sum()
        print(вес_сумма)

        # merging dataframes---------------------------------------------------------------------------------------------------------------------------------------------------
        try:
                df_starka = pd.merge(df_starka, df_поступление_starka, how = "outer", on = ["д_сдачи", "ОП", "корп"])
                df_starka = df_starka.sort_values(by=["д_сдачи", "ОП", "корп"], ascending=True)
        except KeyError:
                df_starka = pd.merge(df_starka, df_поступление_starka, how = "outer", on = ["д_сдачи", "ОП"])
                df_starka = df_starka.sort_values(by=["д_сдачи", "ОП"], ascending=True)
        df_starka["НО-1С_гол"] = df_starka["гол"] - df_starka["ттн_гол"]
        df_starka["НО-1С_вес"] = df_starka["вес"] - df_starka["ттн_вес"]
        df_starka = df_starka.drop(df_starka[(df_starka["НО-1С_гол"] == 0) & (df_starka["НО-1С_вес"] == 0)].index)
        # df_starka.loc[df_starka["НО-1С_вес"] < 0, ["НО-1С_вес"]] = df_starka["НО-1С_вес"]*(-1) # зачем я умножаю?
        df_starka = df_starka.drop(df_starka[(df_starka["НО-1С_гол"] == 0) & (df_starka["НО-1С_вес"] < 0.000000001)].index)
        функции.print_line("hyphens")
        print("\nСРАВНЕНИЕ")
        if df_starka.empty == True:
                print("ВСЕ СХОДИТСЯ")
        if df_starka.empty == False:
                print(df_starka)
