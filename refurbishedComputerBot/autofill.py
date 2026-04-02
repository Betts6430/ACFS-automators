from tkinter import messagebox

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.keys import Keys

from time import sleep
import threading
import sys


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


def show_message(root, title, message):
    """Thread-safe blocking messagebox. Can be called from any thread."""
    done = threading.Event()
    def _show():
        root.wm_attributes("-topmost", 1)
        messagebox.showinfo(title, message)
        root.wm_attributes("-topmost", 0)
        done.set()
    root.after(0, _show)
    done.wait()


def show_warning(root, title, message):
    """Thread-safe blocking warning messagebox."""
    done = threading.Event()
    def _show():
        messagebox.showwarning(title, message)
        done.set()
    root.after(0, _show)
    done.wait()


def enter_orders(data_list, driver, wait, root):
    driver.get("http://173.183.250.6:5014/OrderPages/OrderList.aspx")
    show_message(root, "Action Required", "Please click the order you would like to enter the computers into.")

    try:
        print("Order page detected... Pasting computer barcodes.")

        barcode_field = wait.until(EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_tbx_barcode")))

        # Paste in each value
        for computer in data_list:
            barcode_field.clear()
            barcode_field.send_keys(computer["Computer Barcode"] + Keys.ENTER)
            sleep(0.1)
    
        # Click update order button
        sleep(0.3)
        update_order_btn = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btn_updateOrder")))
        update_order_btn.click()

    except Exception as e:
        print(f"Error entering orders: {e}")
        show_warning(root, "Error", f"Failed to enter orders: {e}")


def open_page(driver, root):
    # Navigate to the home page
    driver.get("http://173.183.250.6:5014/MainPage.aspx")
    show_message(root, "Action Required", "Please login and navigate to the refurbished computers page.")

    # Wait up to 5 minutes for the user to get to the refurbished computers page
    # If the page is reached, the Number of Hard Drives box will be detected
    try:
        wait = WebDriverWait(driver, 300)
        wait.until(EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_tbx_NumOfHardDrives")))
        print("Login detected! Starting the bot...")
        return wait

    except Exception as e:
        print(f"Login timed out: {e}")
        show_warning(root, "Timeout", "Login timed out.")
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


def run_automation(data_list, root_window, order_entry, status_label=None):
    print(f"Processing {len(data_list)} rows...")

    # Open up webdriver page
    driver = webdriver.Chrome()
    if status_label:
        root_window.after(0, lambda: status_label.config(text="Browser is open"))
    wait = open_page(driver, root_window)

    # Cancel program if page is not opened properly
    if wait is None:
        show_warning(root_window, "Timeout", "Login page not reached. Closing.")
        driver.quit()
        return

    try:
        for index, row in enumerate(data_list):
            print(f"Filling row {index + 1}/{len(data_list)}: {row}")

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

    except WebDriverException:
        print("Browser window closed manually.")
        show_warning(root_window, "Cancelled", "Browser window closed manually.")
    except Exception as e:
        print(f"Error occurred: {e}")
        show_warning(root_window, "Error", f"An error occurred: {e}")
    
    else:
        # Only enter orders and show success if automation completed without errors
        if order_entry:
            enter_orders(data_list, driver, wait, root_window)
        show_message(root_window, "Success", f"Finished entering {len(data_list)} computers!")

    finally:
        driver.quit()
        try:
            root_window.after(0, root_window.destroy)
        except Exception:
            pass