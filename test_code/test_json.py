# -*- coding: utf-8 -*-
#!/usr/bin/python
import os
import sys
import json

json_string = '''{
    "id": 1,
    "username": "Bret",
    "email": "Sincere@april.biz",
    "address": {
        "street": "Kulas Light",
        "suite": "Apt. 556",
        "city": "Gwenborough",
        "zipcode": "92998-3874"
    },
    "admin": false,
    "hobbies": null,
    "code_list":{"ga":"경안","gj":"광주","op":"오포"}
}'''


json_object = json.loads(json_string)

assert json_object['id'] == 1
assert json_object['email'] == 'Sincere@april.biz'
assert json_object['address']['zipcode'] == '92998-3874'
assert json_object['admin'] is False
assert json_object['hobbies'] is None

code_list = json_object["code_list"]

print("parse_json result: %s" % type(code_list))

code_key = list[code_list.keys()]
code_value = list[code_list.values()]

print([key for key in code_list])
print([key for key in code_list.keys()])
print([value for value in code_list.values()])
