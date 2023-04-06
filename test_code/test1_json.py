# -*- coding: utf-8 -*-
#!/usr/bin/python
import os
import sys
import json

json_string = '''{
    "stp_list":{"경안":"ga","광주":"gj","오포":"op"},
    "ga": {
		"code": "ga-stp",
		"name": "오포맑은물복원센터",
        "report_type" : {"운영":"op", "고도":"dbf"},

        "op" :{
            "name": "운영보고서",
            "path": "운영",
            "cycle": {"월보":"m", "년보":"y"},
            "daily":{
                "code": "m",
                "name": "오포고도일일보고서",
                "template": "dbf_rptday.xlsx",
                "pg_head_size_in_shts": [],
                "pg_tail_size_in_shts": [],
                "date_cell_in_shts":  [],
                "first_data_cell_in_shts":  [],
                "first_sum_cell_in_shts":  []
            },
            "monthly":{
                "name": "오포고도월간보고서",
                "template": "dbf_rptmonth.xlsx",
                "pg_head_size_in_shts": [],
                "pg_tail_size_in_shts": [],
                "date_cell_in_shts":  [],
                "first_data_cell_in_shts":  [],
                "first_sum_cell_in_shts":  []
            },
            "yearly":{
                "name": "오포고도년간보고서",
                "template": "dbf_rptyear.xlsx",
                "pg_head_size_in_shts": [],
                "pg_tail_size_in_shts": [],
                "date_cell_in_shts":  [],
                "first_data_cell_in_shts":  [],
                "first_sum_cell_in_shts":  []
            }
        },
        
        "dbf" :{
            "name": "운영보고서",
            "path": "운영",
            "cycle": {"월보":"m", "년보":"y"}
        }
    },

    "gj": {
		"code": "op-stp",
		"name": "광주맑은물복원센터",
        "report": {
            "op" :{
                "name": "운영보고서",
                "path": "운영",
                "cycle": {"월보":"m", "년보":"y"}
            },
            "dbf" :{
                "name": "운영보고서",
                "path": "운영",
                "cycle": {"월보":"m", "년보":"y"}
            }
        }
    },

    "op": {
		"code": "op-stp",
		"name": "오포맑은물복원센터",
        "report": {
            "op" :{
                "name": "운영보고서",
                "path": "운영",
                "cycle": {"월보":"m", "년보":"y"}
            },
            "dbf" :{
                "name": "운영보고서",
                "path": "운영",
                "cycle": {"월보":"m", "년보":"y"}
            }
        }
    }
}'''


conf = json.loads(json_string)


code_list = conf["stp_list"]

print("parse_json result: %s" % type(code_list))

stps = [stp for stp in conf["stp_list"].keys()]
rpt_types = [rpt_type for rpt_type in conf["ga"]["report_type"].keys()]
rpt_cycle = [cycle for cycle in conf["ga"]["op"]["cycle"].keys()]
print(conf["ga"]["report_type"], type(conf["ga"]["report_type"]))

rpt = {}
for key, value in conf["ga"]["report_type"].items():
    # rpt[rpt_type] = [cycle for cycle in conf["ga"]["op"]["cycle"].keys()]
    rpt[key] = list(conf["ga"][value]["cycle"].keys())


print(stps, rpt_types, rpt_cycle)