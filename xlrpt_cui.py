#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
    cimon일보를 가져와 월보와 년보를 생성하는 프로그램.
    Usage:
        # Command line run
        # C:>start /b python xlrpt.py -p ga -s 2023-02-01 -e 2023-02-28 -t m
"""

import argparse
import datetime

import xlrpt_utils


def get_arg(**kwargs):
    """
    main 함수
    """
    # rpt_dat: dict = kwargs["rpt_dat"]
    rpt_dat_code: dict = kwargs["rpt_dat_code"]
    # help_msg: str = kwargs["help_msg"]

    stp_code_list: dict = list(rpt_dat_code.keys())

    parser = argparse.ArgumentParser(description="월/년 보고서 생성기 v0.1")
    parser.add_argument(
        "-p",
        "--stp",
        type=str,
        required=True,
        default=None,
        metavar="stp code",
        help=f"The code of the sewage treatment plant for which you want to generate the report must be from "
        f"{list(rpt_dat_code.keys())}.",
    )

    parser.add_argument(
        "-t",
        "--report_type",
        type=str,
        required=True,
        default=None,
        metavar="report type code",
        help=f"Report Generation Type Code" f"{list(rpt_dat_code.keys())}.",  # 메세지 보강
    )

    parser.add_argument(
        "-c",
        "--report_cycle",
        type=str,
        required=True,
        default=None,
        metavar="report cycle code",
        help=f"Report generating cycle code" f"{list(rpt_dat_code.keys())}.",  # 메세지 보강
    )

    parser.add_argument(
        "-s",
        "--start_date",
        type=lambda s: datetime.datetime.strptime(s, "%Y-%m-%d").date(),
        required=False,
        default=datetime.date.today(),
        metavar="start_date",
        help="The report generation start date, defaults to Today",
    )

    parser.add_argument(
        "-e",
        "--end_date",
        type=lambda s: datetime.datetime.strptime(s, "%Y-%m-%d").date(),
        required=False,
        default=None,
        metavar="end_date",
        help="Report generation end date, defaults to the last day of the current month",
    )

    args = parser.parse_args()

    # 명령행 인수 논리 오류 검증 및 명칭 변경

    today: datetime.date = datetime.today()

    if args.stp not in stp_code_list:
        print(f"시설명이 {stp_code_list} 중에 있어야 합니다.")
        args = None
    else:
        report_type_code_list: list = list(rpt_dat_code[args.stp].keys())
        if args.report_type not in report_type_code_list:
            print(f"{args.stp}의 보고서 종별 코드는 {report_type_code_list} 중에 있어야 합니다.")
            args = None
        else:
            report_cycle_code_list: list = rpt_dat_code[args.stp][args.report_type]
            if args.report_type not in report_cycle_code_list:
                print(f"{args.stp}의 {args.report_type} 보고서의 주기 코드는 {report_cycle_code_list} 중에 있어야 합니다.")
                args = None

    if args.start_date is None:
        args.start_date = datetime.date(today.year, today.month, 1)
    elif args.end_date is None:
        args.end_date = xlrpt_utils.last_day_of_month(args.start_date)
    elif args.start_date > args.end_date:
        print("시작일이 종료일보다 클 수 없습니다")
        args = None

    return args
    # return (args.stp, args.start_date, args.end_date, args.month_report, args.year_report)


if __name__ == "__main__":
    import sys

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

    # print("stp code is [" + ", ".join(f"{stp}" for stp in rpt_dat_code.keys()) + "]")

    # print(
    #     f'{list(rpt_dat_code)[0]} {list(rpt_dat_code["op"])[0]} report cycle code is  ['
    #     + ", ".join(f"{cycle}" for cycle in rpt_dat_code["op"]["op"])
    #     + "]"
    # )

    help_msg: str = ""
    for stp_code in rpt_dat_code.keys():
        help_msg = help_msg + stp_code + " ["
        for rpt_code in list(rpt_dat_code[stp_code].keys())[:-1]:
            help_msg = (
                help_msg
                + rpt_code
                + " <"
                + ", ".join(f"{cyc_code}" for cyc_code in rpt_dat_code[stp_code][rpt_code])
                + ">, "
            )
            help_msg = (
                help_msg
                + rpt_code
                + " <"
                + ", ".join(
                    f"{cyc_code}" for cyc_code in rpt_dat_code[stp_code][list(rpt_dat_code[stp_code].keys())[-1]]
                )
                + ">"
            )
        help_msg = help_msg + "] \n"

    # print(cli_help_msg)

    cli_args = get_arg(rpt_dat=rpt_dat, rpt_dat_code=rpt_dat_code, help_msg=help_msg)

    if cli_args is None:
        sys.exit()

    else:
        rpt_para = {
            "stp": cli_args.stp_code,
            "rpt_type": cli_args.rpt_type_code,
            "rpt_cycle": cli_args.rpt_cycle_code,
            "start_date": cli_args.start_date,
            "end_date": cli_args.end_date,
        }

    print(rpt_para)
