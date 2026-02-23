import csv
import urllib.request
import re


def parse_url(url):
    sheet_id_match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
    sheet_id = sheet_id_match.group(1) if sheet_id_match else None
    
    gid_match = re.search(r'gid=([0-9]+)', url)
    gid = gid_match.group(1) if gid_match else "0"
    
    return sheet_id, gid


def get_sheet_data(url, start_row, num_rows):
    sheet_id, gid = parse_url(url)
    if not sheet_id or not gid:
        print("Error")
        return None

    header_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&gid={gid}&range=A1:P1"
    with urllib.request.urlopen(header_url) as f:
        header_line = f.read().decode('utf-8').splitlines()
        headers = next(csv.reader(header_line))
    

    data_range = f"A{start_row}:P{start_row + num_rows - 1}"
    data_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&gid={gid}&range={data_range}"
    with urllib.request.urlopen(data_url) as f:
        data_lines = f.read().decode('utf-8').splitlines()
        data_list = list(csv.DictReader(data_lines, fieldnames=headers))

    # Validate data
    # Check if any serial numbers are NaN. If so, return "Error: Line i read incorrectly."

    return data_list

# DEBUGGING
# data = get_sheet_data('https://docs.google.com/spreadsheets/d/1JN4Cq0aqU9mL1KndfEY4a7F_cSZv3NWILeqjl3UO4uw/edit?gid=292262298#gid=292262298', 230, 5)
# print(data)

