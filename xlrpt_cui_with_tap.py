#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
    cimon일보를 가져와 월보와 년보를 생성하는 프로그램.
    Usage:
        # Command line run
        # C:>start /b python xlrpt.py -p ga -s 2023-02-01 -e 2023-02-28 -t m
"""

from tap import Tap
import datetime
import xlrpt_utils


class XlrptArgPaser(Tap):
    stp: str  # The code of the sewage treatment plant for which you want to generate the report must be from
    rpt_type: str  # Report Generation Type Code
    rpt_cycle: str  # Report generating cycle code
    start_date: str  # The report generation start date, defaults to Today
    end_date: str  # Report generation end date, defaults to the last day of month of the start date


if __name__ == "__main__":
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

    args = XlrptArgPaser().parse_args()

    rpt_para = {
        "stp": args.stp,
        "rpt_type": args.rpt_type,
        "rpt_cycle": args.rpt_cycle,
        "start_date": args.start_date,
        "end_date": args.end_date,
    }

    print(rpt_para)
