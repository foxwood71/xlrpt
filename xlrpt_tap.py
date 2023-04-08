#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
    cimon일보를 가져와 월보와 년보를 생성하는 프로그램.
    Usage:
        # Command line run
        # C:>start /b python xlrpt.py -p ga -s 2023-02-01 -e 2023-02-28 -t m
"""

import datetime

from tap import Tap

import xlrpt_utils


class Cli(Tap):
    """
    Excel report generator v1.0
    """

    stp: str  # Report generation stp code
    report_type: str  # Report generation type Code
    report_cycle: str  # Report generating cycle code
    start_date: datetime.date  # The report generation start date, defaults to Today
    end_date: datetime.date  # Report generation end date, defaults to the last day of month of the start date

    def configure(self) -> None:
        self.add_argument("-p", "--stp", metavar="stp_code")
        self.add_argument(
            "-t",
            "--report_type",
            metavar="report_type_code",
        )
        self.add_argument(
            "-c",
            "--report_cycle",
            metavar="report_cycle_code",
        )
        self.add_argument(
            "-s",
            "--start_date",
            metavar="start_date",
            type=lambda s: datetime.datetime.strptime(s, "%Y-%m-%d").date(),
            default=None,
            required=False,
        )
        self.add_argument(
            "-e",
            "--end_date",
            metavar="end_date",
            type=lambda s: datetime.datetime.strptime(s, "%Y-%m-%d").date(),
            default=None,
            required=False,
        )

    def process_args(self):
        rpt_dat: dict = {}
        rpt_dat_code: dict = {}

        conf: dict = xlrpt_utils.read_config()

        stps: dict = conf["stp_list"]

        for stp_name, stp_code in stps.items():
            rpt_list = {}
            rpt_list_code = {}
            for rpt_name, rpt_code in conf[stp_code]["rpt_type"].items():
                rpt_list[rpt_name] = list(conf[stp_code][rpt_code]["rpt_cycle"].keys())
                rpt_list_code[rpt_code] = list(conf[stp_code][rpt_code]["rpt_cycle"].values())
            rpt_dat_code[stp_code] = rpt_list_code
            rpt_dat[stp_name] = rpt_list

        stp_list_code = list(rpt_dat_code.keys())

        # 명령행 인수 논리 오류 검증 및 명칭 변경
        # 처리시설 코드 확인
        if self.stp in stp_list_code:
            self.report_type
            report_type_code_list: list = list(rpt_dat_code[self.stp].keys())

            # 보고서 코드 확인
            if self.report_type in report_type_code_list:
                report_cycle_code_list: list = rpt_dat_code[self.stp][self.report_type]

                # 보고서 주기 확인
                if self.report_cycle not in report_cycle_code_list:
                    raise ValueError(
                        f"The cycle code of the {self.report_type} report of {self.stp} stp"
                        f"must be in {report_cycle_code_list}."
                    )
            else:
                raise ValueError(f"The report type code of {self.stp} stp must be in {report_type_code_list}.")

        else:
            raise ValueError(f"The stp code must be in {stp_list_code}.")

        # 보고서 작성 시작일 종료일 점검
        if self.end_date is None:
            if self.start_date is None:
                self.start_date = datetime.datetime.today()
            self.end_date = xlrpt_utils.last_day_of_month(self.start_date)

        if self.start_date > self.end_date:
            raise ValueError(
                f"The start date of {self.report_type} report of {self.stp} " f"can't be later than the end date."
            )

    def error(self, message):
        """error(message: string)

        Prints a usage message incorporating the message to stderr and
        exits.

        If you override this in a subclass, it should not return -- it
        should either exit or raise an exception.
        """
        rpt_dat: dict = {}
        rpt_dat_code: dict = {}

        conf: dict = xlrpt_utils.read_config()

        stps: dict = conf["stp_list"]

        for stp_name, stp_code in stps.items():
            rpt_list = {}
            rpt_list_code = {}
            for rpt_name, rpt_code in conf[stp_code]["rpt_type"].items():
                rpt_list[rpt_name] = list(conf[stp_code][rpt_code]["rpt_cycle"].keys())
                rpt_list_code[rpt_code] = list(conf[stp_code][rpt_code]["rpt_cycle"].values())
            rpt_dat_code[stp_code] = rpt_list_code
            rpt_dat[stp_name] = rpt_list

        help_msg: str = ""

        for stp_code in rpt_dat_code.keys():
            help_msg = help_msg + " " * 4 + "stp code : " + stp_code + " --> report code : ["
            for rpt_code in list(rpt_dat_code[stp_code].keys())[:-1]:
                help_msg = (
                    help_msg
                    + rpt_code
                    + " --> report cycle code : <"
                    + ", ".join(f"{cyc_code}" for cyc_code in rpt_dat_code[stp_code][rpt_code])
                    + ">, "
                )
            help_msg = (
                help_msg
                + list(rpt_dat_code[stp_code].keys())[-1]
                + "--> report cycle code : <"
                + ", ".join(
                    f"{cyc_code}" for cyc_code in rpt_dat_code[stp_code][list(rpt_dat_code[stp_code].keys())[-1]]
                )
                + ">"
            )
            help_msg = help_msg + "] \n"

        # print(cli_help_msg)
        message = message + "\n\n - example code list \n\n" + help_msg
        self.print_usage(sys.stderr)
        self.exit(2, ("%s: error msg: %s\n") % (self.prog, message))


if __name__ == "__main__":
    import sys

    try:
        cli_args = Cli().parse_args()

        rpt_para = {
            "stp": cli_args.stp,
            "rpt_type": cli_args.report_type,
            "rpt_cycle": cli_args.report_cycle,
            "start_date": cli_args.start_date,
            "end_date": cli_args.end_date,
        }
        print(rpt_para)

    except ValueError as e:
        print("option error", e)
    except SystemExit:
        exc = sys.exc_info()[1]
        print(f"error code : {exc}")
