#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
    cimon일보를 가져와 월보와 년보를 생성하는 프로그램.
    Usage:
        # Command line run
        # C:>start /b python xlrpt.py -p ga -s 2023-02-01 -e 2023-02-28 -t m
"""

import sys
import warnings

import wmi
from dateutil.relativedelta import relativedelta

import xlrpt_cui
import xlrpt_gui as app_ui
import xlrpt_utils
import xlrpt_xl

cwmi = wmi.WMI()


def main():
    """
    warning message filter
    """
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    """
    main 함수
    """

    conf: dict = xlrpt_utils.read_config()

    rpt_para: dict = {}

    rpt_dat: dict = {}
    rpt_dat_code: dict = {}
    stps: dict = conf["stp_list"]

    for stp_name, stp_code in stps.items():
        rpt_list = {}
        rpt_list_code = {}
        for rpt_name, rpt_code in conf[stp_code]["rpt_type"].items():
            rpt_list[rpt_name] = list(conf[stp_code][rpt_code]["rpt_cycle"].keys())
            rpt_list_code[rpt_code] = list(conf[stp_code][rpt_code]["rpt_cycle"].values())
        rpt_dat_code[stp_code] = rpt_list_code
        rpt_dat[stp_name] = rpt_list

    if len(sys.argv) == 1:

        app_mode_ui = True
        app = app_ui.App(rpt_dat=rpt_dat, rpt_dat_code=rpt_dat_code)
        app.title("보고서생성기 v0.15, 2023.03.29")
        app.mainloop()

        if not app.app_cancel:
            rpt_para = {
                "stp": app.stp_code,
                "rpt_type": app.rpt_type_code,
                "rpt_cycle": app.rpt_cycle_code,
                "start_date": app.start_date,
                "end_date": app.end_date,
            }

        else:
            sys.exit()

        app.destroy()

    else:
        app_mode_ui = False
        cli_args = xlrpt_cui.get_arg(rpt_dat=rpt_dat, rpt_dat_code=rpt_dat_code)

        if cli_args is None:
            sys.exit()

        else:
            rpt_para = {
                "stp": cli_args.stp,
                "start_date": cli_args.start_date,
                "end_date": cli_args.end_date,
                "report_type": cli_args.report_type,
            }

    if rpt_para["stp"] == "전체":
        stps = conf["stp_list"]

    else:
        stps = [
            rpt_para["stp"],
        ]

    diff_date = relativedelta(rpt_para["end_date"], rpt_para["start_date"])  # 두 날짜의 차이 구하기

    diff_months = 12 * diff_date.years + diff_date.months  # 두 날짜의 차이나는 개월수
    diff_years = diff_date.years  # 두 날짜의 차이나는 개월수

    for stp in stps:

        if rpt_para["report_type"] == xlrpt_xl.MONTHLY:  # 월보
            print("---- 월보 ----> " + stp)
            for i in range(0, diff_months + 1, 1):
                rpt_month_frist_date = rpt_para["start_date"].replace(day=1) + relativedelta(months=i)
                rpt_month_last_date = xlrpt_utils.last_day_of_month(rpt_month_frist_date)

                xlrpt_xl.xlsx_rpt(
                    conf=conf,
                    stp=stp,
                    start_date=rpt_month_frist_date,
                    end_date=rpt_month_last_date,
                    type=xlrpt_xl.MONTHLY,
                )

        if rpt_para["report_type"] == xlrpt_xl.YEARLY:  # 월보
            print("---- 년보 ----> " + stp)
            for i in range(0, diff_years + 1, 1):
                rpt_year_first_date = rpt_para["start_date"].replace(month=1, day=1) + relativedelta(years=i)
                rpt_year_last_date = rpt_year_first_date.replace(month=12, day=31)

                xlrpt_xl.xlsx_rpt(
                    conf=conf,
                    stp=stp,
                    start_date=rpt_year_first_date,
                    end_date=rpt_year_last_date,
                    type=xlrpt_xl.YEARLY,
                )

        if rpt_para["report_type"] == xlrpt_xl.FLOW:  # 유량
            print("---- 유량월보 ----> " + stp)
            for i in range(0, diff_months + 1, 1):
                rpt_month_frist_date = rpt_para["start_date"].replace(day=1) + relativedelta(months=i)
                rpt_month_last_date = xlrpt_utils.last_day_of_month(rpt_month_frist_date)

                xlrpt_xl.xlsx_rpt(
                    conf=conf,
                    stp=stp,
                    start_date=rpt_month_frist_date,
                    end_date=rpt_month_last_date,
                    type=xlrpt_xl.FLOW,
                )

    if app_mode_ui:
        app_ui.msgbox("report generating fininsh")
    else:
        print("report generating fininsh")

    sys.exit()


if __name__ == "__main__":
    main()