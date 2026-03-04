from tkinter import messagebox

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

from time import sleep


def select_insensitive(select_element, target_text):
    target = str(target_text).lower().replace(" ", "")

    for option in select_element.options:
        if option.text.lower().replace(" ", "") == target:
            select_element.select_by_visible_text(option.text)
            return True
    return False

def select_ram_insensitive(select_element, target_text):
    target = int(float(str(target_text).lower().replace("gb", "").strip()))

    for option in select_element.options:
        if int(float(option.text.lower().replace("gb", "").strip())) == target:
            select_element.select_by_visible_text(option.text)
            return True
    return False

def select_OS_insensitive(select_element, target_text):
    """Fills in OS dropdown section by checking the input target_text
        Options include Chrome OS, Linux, No OS, Win 10 - MRRP, Win 11 - MRRP."""
    
    target = str(target_text).lower().strip()

    if "no os" in target or "none" in target or not target:
        select_insensitive(select_element, "No OS")
    elif "linux" in target:
        select_insensitive(select_element, "Linux")
    elif "chrome" in target:
        select_insensitive(select_element, "Chrome OS")
    else:
        select_insensitive(select_element, "Win 11 - MRRP")

def fill_in_optical(driver, optical_type):
    """Fills in optical drive dropdown section by checking the input optical_type
        Options include CD x Read, CD x Read/Write, DVD x Read, DVD x Read/Write, Unspecified."""
    
    opticalDrive_dropdown = Select(driver.find_element(By.ID, "ContentPlaceHolder1_ddl_OpticalDrive"))
    target = str(optical_type).lower().strip()

    if "cd" in target:
        if ("rom" in target) or ("read only" in target): opticalDrive_dropdown.select_by_visible_text("CD x Read")
        else: opticalDrive_dropdown.select_by_visible_text("CD x Read/Write")

    elif "dvd" in target:
        if ("rom" in target) or ("read only" in target): opticalDrive_dropdown.select_by_visible_text("DVD x Read")
        else: opticalDrive_dropdown.select_by_visible_text("DVD x Read/Write")

    else:
        opticalDrive_dropdown.select_by_visible_text("Unspecified")



def open_page(driver):
    # Navigate to the home page
    driver.get("http://173.183.250.6:5014/MainPage.aspx")
    messagebox.showinfo("Action Required", "Please login and navigate to the refurbished computers page.")

    # Wait up to 5 minutes for the user to get to the refurbished computers page
    # If the page is reached, the Number of Hard Drives box will be detected
    try:
        wait = WebDriverWait(driver, 300)
        wait.until(EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_tbx_NumOfHardDrives")))
        print("Login detected! Starting the bot...")
        return wait

    except:
        messagebox.showinfo("Login timed out.")
        return None


def fill_page(computer_data, driver, wait):
    # Fill text input fields
    hardDrive_field = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "hdBarcodes")))
    hardDrive_field.clear()
    hardDrive_field.send_keys(computer_data["Hard Drive"])

    computerBarcode_field = driver.find_element(By.ID, "ContentPlaceHolder1_tbx_barcode")
    computerBarcode_field.clear()
    computerBarcode_field.send_keys(computer_data["Computer Barcode"])

    CPU_dropdown = Select(driver.find_element(By.ID, "ContentPlaceHolder1_ddl_CPU"))
    select_insensitive(CPU_dropdown, computer_data["i Series"])

    CPUChipNumber_field = driver.find_element(By.ID, "ContentPlaceHolder1_tbx_ChipNumber")
    CPUChipNumber_field.clear()
    CPUChipNumber_field.send_keys(computer_data["CPU"])

    memory_dropdown = Select(driver.find_element(By.ID, "ContentPlaceHolder1_ddl_memory"))
    select_ram_insensitive(memory_dropdown, computer_data["Total Ram"])

    numOfRam_field = driver.find_element(By.ID, "ContentPlaceHolder1_tbx_NumOfRam")
    numOfRam_field.clear()
    numOfRam_field.send_keys(computer_data["# of RAM"])

    fill_in_optical(driver, computer_data["Optical Drive"])

    OS_dropdown = Select(driver.find_element(By.ID, "ContentPlaceHolder1_ddl_OS"))
    select_OS_insensitive(OS_dropdown, computer_data["OS"])

    newCOA_field = driver.find_element(By.ID, "ContentPlaceHolder1_tbx_NewWindowsCOA")
    newCOA_field.clear()
    if "N/A" not in computer_data["New COA"]:
        newCOA_field.send_keys(computer_data["New COA"])

    print("Form submitted successfully!")


def run_automation(data_list):
    print(f"Processing {len(data_list)} rows...")

    # Open up webdriver page
    driver = webdriver.Chrome()
    wait = open_page(driver)

    # Cancel program if page is not opened properly
    if wait is None:
        messagebox.showwarning("Timeout", "Login page not reached. Closing.")
        driver.quit()
        return

    for index, row in enumerate(data_list):
        print(f"Filling row {index + 1}/{len(data_list)}: {row}")

        try:
            fill_page(row, driver, wait)

            # Wait until user submits refurbished computer and goes to add a new one
            # Then autofill the next computer on the sheet
            new_computer_button = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_Btn_NewComputer")))

            # If it's the last computer we stop here
            if index >= len(data_list) - 1:
                break

            # Otherwise, move to next
            new_computer_button.click()
            
            # Wait for next page to be cleared
            wait.until(lambda d: d.find_element(By.ID, "ContentPlaceHolder1_tbx_barcode").get_attribute("value") == "")
            print("Page cleared. Moving to next computer.")

        except Exception as e:
            messagebox.showerror("Error", f"Stopped at row {index + 1}: {e}")
            break

    # All computers have been entered by this point
    messagebox.showinfo("Success", f"Finished entering {len(data_list)} computers!")
    driver.quit()


# DEBUGGING
# comp_data = {'Serial #': 'PC1WVCJY', 'Hard Drive': 'AB01H-00000081552', 'Computer Barcode': 'AB01C-00000126030', 'Old COA': 'Windows 10 Pro', 'New COA': 3305659803308, 'Initials': 'CP', '#': 2, 'Type': 'Thinkpad X13 Gen 1', 'i Series': 'Xeon Quadcore 1.86GHz', 'CPU': 4650, 'Total Ram': '8GB', '# of RAM': 8, 'Optical Drive': "CD ROM", 'OS': 'Windows 11 Pro', 'Product Key': 'DWV28-BMNJF-QVP96-B2Y84-KBT6P', 'Notes': None}
# fill_page(comp_data)