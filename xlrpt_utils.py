#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
    cimon일보를 가져와 월보와 년보를 생성하는 프로그램.
    Usage:
        # Command line run
        # C:>start /b python xlrpt.py -p ga -s 2023-02-01 -e 2023-02-28 -t m
"""

import datetime
import json
import os
import pathlib
import sys

# print(os.getcwd())
# os.chdir('D:/works/python/xl')

# 폴더 기본세팅
# 기존
# BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
# 3.4 이후
BASE_DIR = pathlib.Path(__file__).resolve().parent.parent


def create_directory(directory):
    """
    main 함수
    """
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print("Error: Failed to create the directory.")


def read_config():
    """
    main 함수
    """
    configuration_file_path = BASE_DIR / "etc/xlrpt_conf.json"
    try:
        with open(configuration_file_path, "r", encoding="utf-8") as file:
            configuration_data = json.load(file)

        return configuration_data

    except FileNotFoundError:
        print("[xlrpt error] - configuration file not found")
        sys.exit(0)


def last_day_of_month(any_day: datetime.date) -> datetime.date:
    """
    main 함수
    """
    # last_day_of_month(datetime.date(2012, month, 1))
    next_month = any_day.replace(day=28) + datetime.timedelta(days=4)  # this will never fail
    return next_month - datetime.timedelta(days=next_month.day)


# loaded excel all close without saving
def close_all_excel_app():
    try:
        os.system("TASKKILL /F /IM excel.exe")

    except Exception:
        print("KU")


if __name__ == "__main__":

    # read configuration file
    xconf = read_config()
    print(len(xconf["경안"]["sum_cell_type"][0]))
    print(len(xconf["경안"]["sum_cell_type"][0][0]))
    print(xconf["경안"]["sum_cell_type"][0][0])
