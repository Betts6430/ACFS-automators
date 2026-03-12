import tkinter as tk
from tkinter import ttk
from google.oauth2 import service_account
from googleapiclient.discovery import build
import subprocess
import os
import tempfile
import sys
import json
from dotenv import load_dotenv

load_dotenv()
key_data = os.getenv('json_creds')#json.loads(os.environ['json_creds'])
creds_dict = json.loads(key_data)


global new_id
new_id= None

def copy_doc(drive_service, TEMPLATE_ID):
    copied_file = drive_service.files().copy(
        fileId=TEMPLATE_ID,
        body={"name": "Temp Order Sheet", "parents": ['1yKEDZSvA3hq8Ci_yYRGzaZgpWRgjd95Y']},
        supportsAllDrives = True,
    ).execute()

    new_doc_id = copied_file["id"]
    return new_doc_id


def export_pdf(drive_service, new_doc_id):
    pdf_request = drive_service.files().export_media(
        fileId=new_doc_id,
        mimeType="application/pdf"
    )

    pdf_data = pdf_request.execute()

    temp_path = os.path.join(tempfile.gettempdir(), "order_label.pdf")

    with open(temp_path, "wb") as f:
        f.write(pdf_data)

    adobe_path = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"

    subprocess.Popen([
        adobe_path,
        "/t",
        temp_path
    ])


def delete_copy(drive_service, NEW_TEMPLATE_ID):
    drive_service.files().delete(
        fileId=NEW_TEMPLATE_ID,
        supportsAllDrives=True,
        enforceSingleParent=True
    ).execute()


def script_init(create_new_copy):
    # setup
    SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
    #SERVICE_ACCOUNT_FILE = json.loads(json_creds)#get_resource_path(key_data)
    TEMPLATE_ID = '1A9jLX_9FuJDY4N2TOMyf_rIZlcpeE4Yk6SE-l-QslaY'

    # authentication
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)

    # run
    docs_service = build('docs', 'v1', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)

    # create copy of template and get the new doc id
    if create_new_copy == True:
        TEMPLATE_ID = copy_doc(drive_service, TEMPLATE_ID)
        global new_id
        new_id = TEMPLATE_ID

    # get doc content
    if new_id == None:
        document = docs_service.documents().get(documentId=TEMPLATE_ID).execute()
    else:
        document = docs_service.documents().get(documentId=new_id).execute()
    content = document["body"]["content"]

    print("Authentication Successful")
    return docs_service, TEMPLATE_ID, drive_service, content


def send(inputs, input_field_box, window, labels, packing_info, current_box):
    for i in range(len(inputs)):
        inputs[i] = input_field_box[i].get()
    print(inputs)

    send_packing_info = [packing_info[0].get(), packing_info[1].get(),packing_info[2].get()]
    print(send_packing_info)

    window.destroy()
    test_script(inputs,labels, send_packing_info, current_box)
    return inputs

def window(order_num_placeholder, org_name_placeholder, total_box_pallet_placeholder, current_box):
    # create window
    root = tk.Tk()
    root.title("Select Contents")

    labels = ['Order Number E-', 'Org. Name' ,'Laptops', 'Desktops', 'Monitors', 'Printers', 'Keyboards', 'Mice', 'Power Cables', 'Display Cables', 'HDMI Cables']
    #global inputs
    inputs = [0,0,0,0,0,0,0,0,0,0,0]
    input_field_box = list(range(len(inputs)))

    j=0
    for i in range(len(inputs)):
        if i == 1:
            quantity_var = tk.StringVar(value=org_name_placeholder)
        else:
            quantity_var = tk.IntVar(value = order_num_placeholder * (i==0))

        if i == 2:
            division = ttk.Separator(root, orient='horizontal')
            division.grid(row=i, column=0, columnspan = 4, sticky='ew', pady=10)
            j=1

        # create label
        label = tk.Label(root, text=labels[i])
        label.grid(row=i + j, column=0, padx=5 * (not i==0), pady=5, stick="e")

        # create input field
        if i != 1:
            input_field_box[i] = tk.Spinbox(root, from_=0, to=30+((i==0) * 10000), textvariable=quantity_var ,width=5 + (i<2)*5 +(i==1)*10)
        else:
            input_field_box[i] = tk.Entry(root, textvariable=quantity_var)
        input_field_box[i].grid(row=i + j, column=1, padx=1, pady=5)

    packing_info = get_box_pallet_info(root,total_box_pallet_placeholder, current_box)

    tk.Button(root, text="Send", command=lambda: send(inputs, input_field_box, root, labels, packing_info, current_box)).grid(row=7, column=2, padx=5, pady=5)
    division = ttk.Separator(root, orient='horizontal')
    division.grid(row=13, column=0, columnspan=4, sticky='ew', pady=5)

    root.mainloop()

def get_box_pallet_info(root, total_box_no_placeholder, current_box):
    # Box, initial value, of, and final value code
    box_pallet_text = tk.StringVar(value='Box')
    box_pallet_box = tk.OptionMenu(root, box_pallet_text, 'Box', 'Pallet')
    box_pallet_box.grid(row=14, column=0, padx=5, pady=5)

    quantity_var0 = tk.IntVar(value=current_box)
    current_box_no = tk.Spinbox(root, from_=0, to=30, textvariable=quantity_var0, width=5)
    current_box_no.grid(row=14, column=1, padx=0, pady=5)

    of = tk.Label(root, text="of")
    of.grid(row=14, column=2, padx=0, pady=5)

    quantity_var = tk.IntVar(value=total_box_no_placeholder)
    total_box_no = tk.Spinbox(root, from_=0, to=30, textvariable=quantity_var, width=5)
    total_box_no.grid(row=14, column=3, padx=5, pady=5)

    return box_pallet_text, current_box_no, total_box_no

def append_replace_req(place_holder, target, requests):
    requests.append(
        {
            "replaceAllText": {
                "containsText": {
                    "text": place_holder,
                    "matchCase": True
                },
                "replaceText": target
            }
        }
    )
    return requests


def find_table(content):
    table = None
    table_start_index = None

    for element in content:
        if "table" in element:
            table = element["table"]
            table_start_index = element["startIndex"]
            break

    if table is None:
        print("No table found.")

    return table, table_start_index


def delete_rows(table, table_start_index, requests):
    # first find the rows to be deleted
    rows_to_delete = []

    for i, row in enumerate(table["tableRows"]):
        first_cell = table["tableRows"][i]
        second_cell = row["tableCells"][1]  # second column
        text = ""

        for content_item in second_cell["content"]:
            if "paragraph" in content_item:
                for element in content_item["paragraph"]["elements"]:
                    if "textRun" in element:
                        text += element["textRun"]["content"]

        if "0x" in text:
            rows_to_delete.append(i)
    print(rows_to_delete)

    #then, remove them from the doc
    for row_index in sorted(rows_to_delete, reverse=True):
        requests.append({
            "deleteTableRow": {
                "tableCellLocation": {
                    "tableStartLocation": {
                        "index": table_start_index
                    },
                    "rowIndex": row_index
                }
            }
        })

    return requests


def exec_req(requests):
    # get table and populate table
    docs_service, NEW_TEMPLATE_ID, drive_service, content = script_init(True)

    docs_service.documents().batchUpdate(
        documentId=NEW_TEMPLATE_ID,
        body={'requests': requests},
    ).execute()

    # delete rows with 0 items
    docs_service, TEMPLATE_ID, drive_service, content = script_init(False)
    table, table_start_index = find_table(content)
    new_requests = delete_rows(table, table_start_index, [])

    if new_requests != []:
        docs_service.documents().batchUpdate(
            documentId=NEW_TEMPLATE_ID,
            body={'requests': new_requests},
        ).execute()

    # create printable pdf
    export_pdf(drive_service, NEW_TEMPLATE_ID)
    delete_copy(drive_service, NEW_TEMPLATE_ID)


def test_script(inputs, labels, packing_info, current_box):
    requests = []
    packing_labels = ['BP','Current','Total']
    i=0
    for input in inputs:
        if input == '0':
            input += 'x'
        append_replace_req('x' + labels[i], input, requests)
        i+=1
    j=0
    for info in packing_info:
        append_replace_req('x' + packing_labels[j], info, requests)
        j+=1

    exec_req(append_replace_req('Izaan','working?', requests))
    new_win_dialog(inputs, packing_info, current_box)


##################### new window creation #######################
def yes_or_no(win, result, inputs, packing_info, current_box):
    x,x,total = packing_info
    print(result)
    win.destroy()
    if result:
        print(inputs[0],inputs[1], total)
        window(inputs[0], inputs[1], total, current_box + 1)  # ordernum, org name

def new_win_dialog(inputs, packing_info, current_box):
    win = tk.Tk()
    win.title("Continue?")

    cont_text = tk.Label(win, text="Would you like to create another order sheet?")
    cont_text.pack(pady=10, padx=10)

    yes = tk.Button(win, text="Yes", command=lambda: yes_or_no(win,True, inputs, packing_info, current_box)).pack(pady=5, padx=10)
    no = tk.Button(win, text="No", command=lambda: yes_or_no(win,False, inputs, packing_info, current_box)).pack(pady=5, padx=10)

    win.mainloop()

order_num = 6767
org_name = 'Name'

window(order_num,org_name,1,1)