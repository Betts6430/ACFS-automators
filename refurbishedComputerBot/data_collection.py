import pandas as pd
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

    header_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&gid={gid}&range=A1:Z1"
    headers = pd.read_csv(header_url, usecols=range(16)).columns

    data_range = f"A{start_row}:P{start_row + num_rows - 1}"
    data_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&gid={gid}&range={data_range}"
    df = pd.read_csv(data_url, header=None, usecols=range(16))
    
    df.columns = headers # Reattach header
    data_list = df.to_dict(orient='records')

    # Validate data
    # Check if any serial numbers are NaN. If so, return "Error: Line i read incorrectly."

    return data_list


# data = get_sheet_data('https://docs.google.com/spreadsheets/d/1UeIMdClT0weM32UDPfFpa_-9gNAb6vnJ33Cu9ivrfwU/edit?gid=0#gid=0', 5, 5)
# print(data)

