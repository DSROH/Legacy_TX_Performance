from io import StringIO
import os, sys, subprocess
import psutil

import numpy as np
import pandas as pd

pd.set_option("display.max_rows", None)
pd.set_option("display.width", None)
pd.set_option("display.max_columns", None)
pd.set_option("display.max_colwidth", None)

import xlwings as xw

import matplotlib.pyplot as plt

plt.rcParams.update({"figure.max_open_warning": 0})

from matplotlib.backends.backend_pdf import PdfPages

import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox as msg


def return_print(*prt_str):
    io = StringIO()
    print(*prt_str, file=io, sep=",", end="")
    return io.getvalue()


def isNaN(num):
    return num != num


def save_multi_image(f_name):
    pp = PdfPages(f_name)
    fig_nums = plt.get_fignums()
    figs = [plt.figure(n) for n in fig_nums]
    for fig in figs:
        fig.savefig(pp, format="pdf")
    pp.close()


def open_file(filename):
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])


def add_file(Entry_file_path):
    Entry_file_path.delete(0, tk.END)
    filename = filedialog.askopenfilenames(
        initialdir="C:\Labtest\Report\Spring",
        title="Select file",
        filetypes=(("All fiels", "*.*"), ("Excel files", "*.xlsx")),
    )
    Entry_file_path.insert(tk.END, filename)


def ET_perf_drawing(Win_GUI, Entry_file_path, text_area):
    try:
        filename = Entry_file_path.get()
        if filename:
            f_name = f"{os.path.splitext(filename)[0]}.pdf"  # filename을 확장자를 지운 후 pdf 확장자로 지정
            TestItem = []
            text_area.insert(tk.END, f"Open Excel File ... \n")
            # openpyxl
            # workbook = xl.load_workbook(r"{}".format(filename))
            # TestItem = [sheet.title for sheet in workbook.worksheets if sheet.sheet_state == "visible"]
            # xlwings
            app = xw.App(visible=False)
            wb = app.books.open(filename)
            # You need to install xlrd==1.2.0 to get the support for xlsx excel format.
            TestItem = [sheet.name for sheet in wb.sheets if sheet.api.Visible == -1]

        dict_value = ["HSDPA", "HSUPA", "WCDMA NORMAL", "WCDMA ALL CHANNEL", "LTE"]
        Plot_list = [s for s in TestItem if any(xs in s for xs in dict_value)]

        for c, i in enumerate(Plot_list, start=1):
            text_area.insert(tk.END, f"Item count {c:<17}|    {i:<20}\n")
            text_area.see(tk.END)
        # list로 받으면 Dataframe 동적변수 선언 시 공백이 포함되서 에러 발생함. -> Str 변환 -> 자리수 자르기 -> List 변환
        # Plot_list=",".join(Plot_list).replace(' ALL CH', '').split(',')                # LTE Normal로 측정 시 사용 // 필요없어 진 듯...
        text_area.insert(tk.END, f"Loading Workbook Done\n")
        text_area.insert(tk.END, "=" * 73)
        text_area.insert(tk.END, "\n")
        text_area.see(tk.END)
        # Plot_list에 있는 시트만 선택하여 데이터프레임 생성하기
        # df_Data = pd.ExcelFile(filename).parse(sheet_name=Plot_list)  # openpyxl
        # df_Data = pd.read_excel(open(wb.fullname, 'rb'), sheet_name=Plot_list, engine='openpyxl') # xlwings -> openpyxl code 지만 제대로 동작 하지 않음
        df_Data = {}
        for sheet_name in Plot_list:
            sheet = wb.sheets[sheet_name]
            df_Data[sheet_name] = sheet.used_range.options(pd.DataFrame, index=False).value
        # 엑셀 프로그램 종료
        wb.close()
        app = xw.apps.active
        for proc in psutil.process_iter():
            if proc.name() == "EXCEL.EXE":
                proc.kill()

        for i in range(len(Plot_list)):

            if any("LTE" in c for c in Plot_list):
                text_area.insert(tk.END, f"Drawing {Plot_list[i]:<20}|    ")
                df_Band = pd.DataFrame(df_Data[Plot_list[i]])
                df_Band = df_Band.drop(index=[0, 1, 2, 3, 4, 5, 6, 7]).reset_index(drop=True)
                df_Band.columns = df_Band.iloc[0]
                Plot_list[i] = Plot_list[i].replace(" ALL CH", "")

                df_BW = (
                    df_Band["BW"].drop_duplicates(keep="first").reset_index(drop=True).dropna().values.tolist()
                )  # 중복값 삭제, NaN Drop, index reset
                df_BW = (return_print(*df_BW)).split(",")  # 각 밴드의 측정 BW 데이터 추출

                fig, ax = plt.subplots(nrows=2, ncols=2, figsize=(30.6, 15.9))
                fig.suptitle(("{} TX Performance".format(Plot_list[i])), fontsize=25)

                ax[0][0].set_title("TX Power")
                ax[0][0].set_ylim(20, 26)  # set_ylim(bottom, top)
                ax[0][0].set_xlabel("Channel", fontsize=10)
                ax[0][0].set_ylabel("TX Power", fontsize=10)
                ax[0][0].grid(True, color="black", alpha=0.3, linestyle="--")

                ax[0][1].set_title("Error Vector Magnitude")
                ax[0][1].set_ylim(0, 5)  # set_ylim(bottom, top)
                ax[0][1].set_xlabel("Channel", fontsize=10)
                ax[0][1].set_ylabel("EVM", fontsize=10)
                ax[0][1].grid(True, color="black", alpha=0.3, linestyle="--")
                ax[0][1].axhline(y=3, linestyle="dashdot", color="red", label="Spec")  # Spec 기준선 Drwaing, EVM 3 Under

                ax[1][0].set_title("Spectrum Emission Mask")
                ax[1][0].set_ylim(-30, -5)  # set_ylim(bottom, top)
                ax[1][0].set_xlabel("Channel", fontsize=10)
                ax[1][0].set_ylabel("SEM", fontsize=10)
                ax[1][0].grid(True, color="black", alpha=0.3, linestyle="--")
                ax[1][0].axhline(
                    y=-6, linestyle="dashdot", color="red", label="Spec"
                )  # Spec 기준선 Drwaing, SEM Margin -6dB

                ax[1][1].set_title("Adjacent Channel Leakage Ratio")
                ax[1][1].set_ylim(-46, -30)  # set_ylim(bottom, top)
                ax[1][1].set_xlabel("Channel", fontsize=10)
                ax[1][1].set_ylabel("ACLR", fontsize=10)
                ax[1][1].grid(True, color="black", alpha=0.3, linestyle="--")
                ax[1][1].axhline(
                    y=-36, linestyle="dashdot", color="red", label="Spec"
                )  # Spec 기준선 Drwaing. ACLR -36dBc 이하 관리

                for Bandwidth in df_BW:

                    Data_of_BW = df_Band[df_Band["BW"] == Bandwidth]
                    Data_of_BW = Data_of_BW.replace(
                        "-", np.nan
                    )  # 미측정으로 인한 오류 (-) 해결을 위해 Nan으로 대체하고 아래에서 dropna 로 제거함.

                    if Bandwidth == "BW":
                        df_CH = Data_of_BW.reset_index(drop=True).iloc[:, 9:]
                        df_CH.index = df_BW[1::1]  # df_BW list slicing [start:stop:step]
                        df_CH = df_CH.transpose().reset_index(drop=True)
                        continue
                    else:
                        CH_List = df_CH[Bandwidth].dropna()
                        df_TXL = (
                            Data_of_BW[Data_of_BW["Test Item"].str.contains("6.2.2 Maximum Output Power_RB")]
                            .iloc[2:3, 9:]
                            .dropna(axis=1)
                            .iloc[0]
                        )
                        df_TXL.index = CH_List
                        df_TXR = (
                            Data_of_BW[Data_of_BW["Test Item"].str.contains("6.2.2 Maximum Output Power_RB")]
                            .iloc[3:4, 9:]
                            .dropna(axis=1)
                            .iloc[0]
                        )
                        df_TXR.index = CH_List
                        df_EVM = (
                            Data_of_BW[Data_of_BW["Test Item"].str.contains("6.5.2.1 EVM")]
                            .iloc[:12, 9:]
                            .dropna(axis=1)
                            .max()
                        )
                        df_EVM.index = CH_List
                        df_SEM = (
                            Data_of_BW[Data_of_BW["Test Item"].str.contains("6.6.2.1 SEM_")]
                            .iloc[:, 9:]
                            .dropna(axis=1)
                            .max()
                        )
                        df_SEM.index = CH_List
                        df_ACLR = 0 - (
                            Data_of_BW[Data_of_BW["Test Item"].str.contains("6.6.2.3 ACLR_")]
                            .iloc[:, 9:]
                            .dropna(axis=1)
                            .min()
                        )
                        df_ACLR.index = CH_List

                        # 특정 subplot을 Twinx 설정 시 ax2=ax[0][1].twinx()
                        ax[0][0].plot(df_TXL, marker=".", label="TX Power {}MHz RB Low".format(Bandwidth))
                        ax[0][0].plot(df_TXR, marker=".", label="TX Power {}MHz RB High".format(Bandwidth))
                        ax[0][0].legend(fontsize=10, frameon=False, loc="lower center", ncol=6)

                        ax[0][1].plot(df_EVM, marker=".", label="EVM {}MHz".format(Bandwidth))
                        ax[0][1].legend(fontsize=12, frameon=False, loc="lower center", ncol=7)

                        ax[1][0].plot(df_SEM, marker=".", label="SEM {}MHz".format(Bandwidth))
                        ax[1][0].legend(fontsize=12, frameon=False, loc="lower center", ncol=7)

                        ax[1][1].plot(df_ACLR, marker=".", label="ACLR {}MHz".format(Bandwidth))
                        ax[1][1].legend(fontsize=12, frameon=False, loc="lower center", ncol=7)

                        plt.tight_layout()

                text_area.insert(tk.END, f"Done\n")
                text_area.see(tk.END)
            else:  # WCDMA
                text_area.insert(tk.END, f"\nDrawing {Plot_list[i]:<20}\n")
                df_Band = pd.DataFrame(df_Data[Plot_list[i]])
                df_Band = df_Band.drop(index=[0, 1, 2, 3, 4, 5, 6]).dropna(how="all", axis=0).reset_index(drop=True)
                df_ItemList = df_Band["Samsung Lab Test Report"]

                list_Range = []
                Band_list = []

                for j, val1 in enumerate(df_ItemList):  # BAND 구분되는 위치를 먼저 확인
                    if "BAND" in val1 or "Band" in val1:
                        list_Range.append(j)
                        Band_list.append(val1)

                Item_Count = list_Range[1] - list_Range[0]

                for k in range(len(list_Range)):  # 데이터 갯수만큼 실행하기 위해 for k in list_Range 대신 len(list_Range) 사용
                    text_area.insert(tk.END, f"Drawing {Band_list[k]:<20}|    ")

                    fig, ax = plt.subplots(nrows=2, ncols=2, figsize=(30.6, 15.9))
                    fig.suptitle(("{} {}".format(Plot_list[i], Band_list[k])), fontsize=25)

                    ax[0][0].set_title("TX Power")
                    ax[0][0].set_xlabel("Channel", fontsize=10)
                    ax[0][0].set_ylabel("TX Power", fontsize=10)
                    ax[0][0].grid(True, color="black", alpha=0.3, linestyle="--")

                    if Plot_list[i] == "HSUPA":
                        ax[0][0].set_ylim(16, 24)
                    else:
                        ax[0][0].set_ylim(20, 26)  # set_ylim(bottom, top)

                    ax[0][1].set_title("Error Vector Magnitude")
                    ax[0][1].set_ylim(0, 5)  # set_ylim(bottom, top)
                    ax[0][1].set_xlabel("Channel", fontsize=10)
                    ax[0][1].set_ylabel("EVM", fontsize=10)
                    ax[0][1].grid(True, color="black", alpha=0.3, linestyle="--")
                    ax[0][1].axhline(
                        y=3, linestyle="dashdot", color="red", label="Spec"
                    )  # Spec 기준선 Drwaing, EVM 3 Under

                    ax[1][0].set_title("Spectrum Emission Mask")
                    ax[1][0].set_ylim(-60, -0)  # set_ylim(bottom, top)
                    ax[1][0].set_xlabel("Channel", fontsize=10)
                    ax[1][0].set_ylabel("SEM", fontsize=10)
                    ax[1][0].grid(True, color="black", alpha=0.3, linestyle="--")
                    ax[1][0].axhline(
                        y=-6, linestyle="dashdot", color="red", label="Spec"
                    )  # Spec 기준선 Drwaing, SEM Margin -6dB

                    ax[1][1].set_title("Adjacent Channel Leakage Ratio")
                    ax[1][1].set_ylim(-50, -32)  # set_ylim(bottom, top)
                    ax[1][1].set_xlabel("Channel", fontsize=10)
                    ax[1][1].set_ylabel("ACLR", fontsize=10)
                    ax[1][1].grid(True, color="black", alpha=0.3, linestyle="--")
                    ax[1][1].axhline(
                        y=-36, linestyle="dashdot", color="red", label="Spec"
                    )  # Spec 기준선 Drwaing. ACLR -36dBc 이하 관리

                    df_TestBand = df_Band.iloc[list_Range[k] : list_Range[k] + Item_Count, :].reset_index(drop=True)
                    # Nan이 행/열 모두 있어서 2번 실행함
                    df_TestBand = df_TestBand.replace("-", np.nan).iloc[1:Item_Count, :].dropna(how="all", axis=1)
                    df_TestBand = df_TestBand.dropna().reset_index(drop=True)
                    df_TestBand.columns = df_TestBand.iloc[0]

                    if Plot_list[i] == "WCDMA NORMAL" or Plot_list[i] == "WCDMA ALL CHANNEL":

                        df_TX_Power = (
                            df_TestBand[df_TestBand["Test Item"].str.contains("5.2 Maximum output power")]
                            .iloc[:, 4:]
                            .dropna(axis=1)
                            .max()
                        )
                        df_SEM = (
                            df_TestBand[df_TestBand["Test Item"].str.contains("5.9 Spectrum emission mask")]
                            .iloc[:, 4:]
                            .dropna(axis=1)
                            .max()
                        )
                        df_ACLR = (
                            df_TestBand[df_TestBand["Test Item"].str.contains("5.10 Adjacent Ch. leakage power ratio")]
                            .iloc[:, 4:]
                            .dropna(axis=1)
                            .max()
                        )
                        df_EVM = (
                            df_TestBand[df_TestBand["Test Item"].str.contains("5.13.1 EVM @ Max Pwr")]
                            .iloc[:, 4:]
                            .dropna(axis=1)
                            .max()
                        )

                        ax[0][0].plot(
                            df_TX_Power, marker=".", label="{} {} TX Max Power".format(Plot_list[i], Band_list[k])
                        )
                        ax[0][0].legend(fontsize=12, frameon=False, loc="lower center", ncol=6)
                        ax[0][1].plot(df_EVM, marker=".", label="{} {} EVM".format(Plot_list[i], Band_list[k]))
                        ax[0][1].legend(fontsize=12, frameon=False, loc="lower center", ncol=6)
                        ax[1][0].plot(df_SEM, marker=".", label="{} {} SEM".format(Plot_list[i], Band_list[k]))
                        ax[1][0].legend(fontsize=12, frameon=False, loc="lower center", ncol=6)
                        ax[1][1].plot(df_ACLR, marker=".", label="{} {} ACLR".format(Plot_list[i], Band_list[k]))
                        ax[1][1].legend(fontsize=12, frameon=False, loc="lower center", ncol=6)

                        plt.tight_layout()

                    else:

                        df_SubList = (
                            df_TestBand["Subtest"]
                            .drop_duplicates(keep="first")
                            .reset_index(drop=True)
                            .dropna()
                            .values.tolist()
                        )  # 중복값 삭제, NaN Drop, index reset
                        df_SubList = (return_print(*df_SubList)).split(",")

                        for SubTest in df_SubList:

                            if SubTest == "Subtest":
                                df_CH = df_TestBand.iloc[:1, 5:].replace("-", np.nan).dropna(axis=1)
                                df_CH = df_CH.transpose().reset_index(drop=True)
                                continue

                            else:
                                # 신규_Dataframe = 추출대상_dataframe[추출대상_dataframe['필터명']==필터의 변수]
                                # 글로벌 변수 공백 시 Variable Explorer 에서 확인 안 됨 -> 공백 제거 SubTest.replace(" " , "")
                                df_SubTest = df_TestBand[df_TestBand["Subtest"] == SubTest].reset_index(drop=True)
                                SubTest = SubTest.replace(" ", "")

                                if Plot_list[i] == "HSDPA":
                                    df_TX_Power = (
                                        df_SubTest[df_SubTest["Test Item"].str.contains("5.2A")]
                                        .iloc[:, 5:]
                                        .dropna(axis=1)
                                        .max()
                                    )
                                    df_SEM = (
                                        df_SubTest[df_SubTest["Test Item"].str.contains("5.9A")]
                                        .iloc[:, 5:]
                                        .dropna(axis=1)
                                        .max()
                                    )
                                    df_ACLR = (
                                        df_SubTest[df_SubTest["Test Item"].str.contains("5.10A")]
                                        .iloc[:, 5:]
                                        .dropna(axis=1)
                                        .max()
                                    )

                                else:
                                    df_TX_Power = (
                                        df_SubTest[df_SubTest["Test Item"].str.contains("5.2B")]
                                        .iloc[:, 5:]
                                        .dropna(axis=1)
                                        .max()
                                    )
                                    df_SEM = (
                                        df_SubTest[df_SubTest["Test Item"].str.contains("5.9B")]
                                        .iloc[:, 5:]
                                        .dropna(axis=1)
                                        .max()
                                    )
                                    df_ACLR = (
                                        df_SubTest[df_SubTest["Test Item"].str.contains("5.10B")]
                                        .iloc[:, 5:]
                                        .dropna(axis=1)
                                        .max()
                                    )

                            ax[0][0].plot(df_TX_Power, marker=".", label="{} TX Max Power".format(SubTest))
                            ax[0][0].legend(fontsize=12, frameon=False, loc="lower center", ncol=6)
                            ax[1][0].plot(df_SEM, marker=".", label="{} SEM".format(SubTest))
                            ax[1][0].legend(fontsize=12, frameon=False, loc="lower center", ncol=6)
                            ax[1][1].plot(df_ACLR, marker=".", label="{} ACLR".format(SubTest))
                            ax[1][1].legend(fontsize=12, frameon=False, loc="lower center", ncol=6)

                            plt.tight_layout()

                    text_area.insert(tk.END, f"Done\n")
                    text_area.see(tk.END)

        text_area.insert(tk.END, f"Saving Image                |   ")
        save_multi_image(f_name)
        text_area.insert(tk.END, f"Done\n")
        text_area.see(tk.END)

        open_file(f_name)

        Win_GUI.destroy()

    except Exception as e:
        msg.showwarning("Warning", e)
