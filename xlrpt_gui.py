#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
    cimon일보를 가져와 월보와 년보를 생성하는 프로그램.
    Usage:
        # Command line run
        # C:>start /b python xlrpt.py -p ga -s 2023-02-01 -e 2023-02-28 -t m
"""

import datetime
import tkinter as tk
import tkinter.messagebox
from tkinter import ttk

from tkcalendar import DateEntry

import xlrpt_utils


def msgbox(message: str) -> None:
    tkinter.messagebox.showinfo("알림", message)


class App(tk.Tk):
    """
    App Class
    """

    __rpt_dat_code: dict

    stp: str  # The code of the sewage treatment plant for which you want to generate the report must be from
    report_type: str  # Report Generation Type Code
    report_cycle: str  # Report generating cycle code
    start_date: datetime.date  # The report generation start date, defaults to Today
    end_date: datetime.date  # Report generation end date, defaults to the last day of month of the start date

    def stp_selected(self, rpt_dat):
        self.cbo_rpt["values"] = list(rpt_dat[self.cbo_stp.get()].keys())
        self.cbo_rpt.current(0)
        self.rpt_selected(rpt_dat)

    def rpt_selected(self, rpt_dat):
        self.cbo_cyc["values"] = list(rpt_dat[self.cbo_stp.get()][self.cbo_rpt.get()])
        self.cbo_cyc.current(0)

    def __init__(self, **kwargs):
        super().__init__()

        # stp_list =  [stp for stp in kwargs['stp_list'].keys()]
        # report_type_list = [rpt_type for rpt_type in kwargs['report_type_list'].keys()]

        rpt_dat = kwargs["rpt_dat"]
        self.__rpt_dat_code = kwargs["rpt_dat_code"]

        stp_list = list(rpt_dat.keys())

        # root window
        self.geometry("260x165")
        self.eval("tk::PlaceWindow . center")
        self.resizable(False, False)
        self.style = ttk.Style()

        # stp select
        self.lb_stp = ttk.Label(self, text="시설명")
        self.lb_stp.grid(row=0, column=0, columnspan=1, padx=5, pady=5)

        self.cbo_stp = ttk.Combobox(self)
        self.cbo_stp.config(width=12, justify=tk.CENTER, values=stp_list, state="readonly")
        self.cbo_stp.option_add("*TCombobox*Listbox.Justify", "center")
        self.cbo_stp.grid(row=0, column=1, columnspan=2, padx=5, pady=5)
        self.cbo_stp.current(0)
        # self.cbo_stp.set(stp_list[0])
        self.cbo_stp.bind("<<ComboboxSelected>>", lambda event: self.stp_selected(rpt_dat))

        # report type select
        self.lb_rpt = ttk.Label(self, text="보고서")
        self.lb_rpt.grid(row=1, column=0, columnspan=1, padx=5, pady=5)

        self.cbo_rpt = ttk.Combobox(self)
        self.cbo_rpt.config(
            width=12, justify=tk.CENTER, values=list(rpt_dat[self.cbo_stp.get()].keys()), state="readonly"
        )
        self.cbo_rpt.option_add("*TCombobox*Listbox.Justify", "center")
        self.cbo_rpt.grid(row=1, column=1, columnspan=2, padx=5, pady=5)
        self.cbo_rpt.current(0)
        # self.cbo_rpt_type.set(report_type_list[0])
        self.cbo_rpt.bind("<<ComboboxSelected>>", lambda event: self.rpt_selected(rpt_dat))

        # report cycle select
        self.lb_cyc = ttk.Label(self, text="주  기")
        self.lb_cyc.grid(row=2, column=0, columnspan=1, padx=5, pady=5)

        self.cbo_cyc = ttk.Combobox(self)
        self.cbo_cyc.config(
            width=12, justify=tk.CENTER, values=list(rpt_dat[self.cbo_stp.get()][self.cbo_rpt.get()]), state="readonly"
        )
        self.cbo_cyc.option_add("*TCombobox*Listbox.Justify", "center")
        self.cbo_cyc.grid(row=2, column=1, columnspan=2, padx=5, pady=5)
        self.cbo_cyc.current(0)
        # self.cbo_rpt_cycle.set(rpt_list[0])

        # DaraEntry -- StartDay
        self.lb_start_date = ttk.Label(self, text="시작일")
        self.lb_start_date.grid(row=3, column=0, columnspan=1, padx=5, pady=5)

        self.txt_start_date = DateEntry(
            self,
            width=12,
            background="darkblue",
            foreground="white",
            borderwidth=2,
            justify=tk.CENTER,
            selectmode="day",
            locale="ko_KR",
            cursor="hand1",
            date_pattern="yyyy-mm-dd",
            firstweekday="sunday",
            showweeknumbers=False,
        )
        self.txt_start_date.set_date(datetime.date.today())
        self.txt_start_date.grid(row=3, column=1, columnspan=2, padx=5, pady=5)

        self.lb_end_date = ttk.Label(self, text="종료일")
        self.lb_end_date.grid(row=4, column=0, columnspan=1, padx=5, pady=5)

        self.txt_end_date = DateEntry(
            self,
            width=12,
            background="darkblue",
            foreground="white",
            borderwidth=2,
            justify=tk.CENTER,
            selectmode="day",
            locale="ko_KR",
            cursor="hand1",
            date_pattern="yyyy-mm-dd",
            firstweekday="sunday",
            showweeknumbers=False,
        )
        self.txt_end_date.set_date(xlrpt_utils.last_day_of_month(datetime.date.today()))
        self.txt_end_date.grid(row=4, column=1, columnspan=2, padx=5, pady=5)

        # button
        btn_ok = ttk.Button(self, text="실행")
        btn_ok.grid(row=0, column=3, padx=0, pady=5, sticky="w")
        btn_ok.bind("<Button-1>", self.cmd_ok_click)  # 사건묶기

        btn_cancel = ttk.Button(self, text="취소")
        btn_cancel.grid(row=1, column=3, padx=0, pady=5, sticky="w")
        btn_cancel.bind("<Button-1>", self.cmd_cancel_click)  # 사건묶기

        self.protocol("WM_DELETE_WINDOW", self.wm_destroyer)

    def cmd_ok_click(self, event):
        """
        main 함수
        """
        self.app_cancel = False
        self.stp_code = list(self.__rpt_dat_code.keys())[self.cbo_stp.current()]  # get() -> value,  current -> index
        self.rpt_type_code = list(self.__rpt_dat_code[self.stp_code].keys())[self.cbo_rpt.current()]
        self.rpt_cycle_code = self.__rpt_dat_code[self.stp_code][self.rpt_type_code][self.cbo_cyc.current()]
        self.start_date = self.txt_start_date.get_date()
        self.end_date = self.txt_end_date.get_date()

        if self.start_date >= self.end_date:
            tkinter.messagebox.showinfo("알림", "시작일이 종료일보다 클 수 없습니다.")
        else:
            # self.destroy()
            self.quit()

    def cmd_cancel_click(self, event):
        """
        main 함수
        """
        self.app_cancel = True
        self.quit()

    def wm_destroyer(self):
        """
        main 함수
        """
        self.app_cancel = True
        self.quit()


if __name__ == "__main__":
    conf = xlrpt_utils.read_config()

    stps = conf["stp_list"]
    rpt_dat: dict = {}
    rpt_dat_code: dict = {}

    for stp_name, stp_code in stps.items():
        rpt_list: dict = {}
        rpt_list_code: dict = {}
        for rpt_name, rpt_code in conf[stp_code]["rpt_type"].items():
            rpt_list[rpt_name] = list(conf[stp_code][rpt_code]["rpt_cycle"].keys())
            rpt_list_code[rpt_code] = list(conf[stp_code][rpt_code]["rpt_cycle"].values())
        rpt_dat_code[stp_code] = rpt_list_code
        rpt_dat[stp_name] = rpt_list

    app = App(rpt_dat=rpt_dat, rpt_dat_code=rpt_dat_code)
    app.title("월/년 보고서 생성기 v0.1")
    app.mainloop()

    # for i, (key, value) in enumerate(my_dict.items()):
    #     (f'Index: {i}, Key: {key}, Value: {value}')

    if not app.app_cancel:
        rpt_para = {
            "stp": app.stp_code,
            "rpt_type": app.rpt_type_code,
            "rpt_cycle": app.rpt_cycle_code,
            "start_date": app.start_date,
            "end_date": app.end_date,
        }
        print(rpt_para)

    app.destroy()
