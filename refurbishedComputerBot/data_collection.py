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

    header_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}&range=A1:P1"
    with urllib.request.urlopen(header_url) as f:
        header_line = f.read().decode('utf-8').splitlines()
        headers = next(csv.reader(header_line))
    

    data_range = f"A{start_row}:P{start_row + num_rows - 1}"
    data_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}&range={data_range}"
    with urllib.request.urlopen(data_url) as f:
        data_lines = f.read().decode('utf-8').splitlines()
        data_list = [{k: v.strip() for k, v in row.items()} for row in csv.DictReader(data_lines, fieldnames=headers, restval='')]

    # Validate data
    # Check if any cells are unfilled on the spreadsheet
    row = 1
    valid_data = True
    error_message = ""

    for comp_data in data_list:
        for category, data in comp_data.items():
            if not data: # Append to error_message all errors instead of just one
                error_message += f"Error: Missing Data - Row {row}, {category}\n"
                valid_data = False
        
        row += 1


    return data_list, valid_data, error_message

# DEBUGGING
# data, valid, _ = get_sheet_data('', 147, 7)
# print(data)
# print(valid)

