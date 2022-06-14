#! /usr/bin/env python3

"""
Usage:
python3 csv_to_json.py <excel-file-name>.xlsx
"""
import json
from pathlib import Path
import sys
import re
from openpyxl import load_workbook

"""Read XLSX file and write it to JSON file in same folder.
"""
def main(xlsx_file):
    workbook = load_workbook(filename=xlsx_file, read_only=True)
    worksheet = workbook['sheet_name']
    excel_data = list(worksheet.rows)
    column_names = [column.value for column in excel_data[0]]

    json_output = []
    for row in excel_data[1:]:
        values = [cell.value for cell in row]
        row_dict = {name: str(value) for name, value in zip(column_names, values)}
        json_output.append(row_dict)

    output = []
    for vul in json_output:
        cve=vul.get('id')
        imgpks=vul.get('package')
        impact=vul.get('impact')
        library=vul.get('lib')
        ids = re.split(', |\n',cve) #Checks for multiple id's delimited by comma or seperate line
        packages = re.split(', |\n',imgpks) #Checks for multiple imgpkgs delimited by comma or seperate line
        for id in ids:
            for package in packages:           
                values = {'id': id, 'package': library, 'repository': package, 'justification': impact}
                output.append(values)

    output_file = Path(xlsx_file).with_suffix(".json")
    with open(output_file, "w") as fhandle:
        fhandle.write(json.dumps(output, indent=4))


if __name__ == "__main__":
    main(sys.argv[1])