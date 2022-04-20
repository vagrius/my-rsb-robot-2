from RPA.Robocorp.Vault import Vault
from RPA.Browser.Selenium import Selenium
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.Dialogs import Dialogs

import csv
import requests
import os
import shutil
import time

browser = Selenium()
pdf = PDF()

def input_dialog():
    text_input = Dialogs()
    text_input.add_heading(heading='Enter the link for CSV-file, please')
    text_input.add_text_input(name='link', label='Link')
    result = text_input.run_dialog()
    return result['link']


def read_from_the_local_vault():
    _secret = Vault().get_secret("data")
    URL = _secret["url"]
    return URL

def get_orders_from_csv(csv_url):    
    # download the CSV-file
    request = requests.get(csv_url)
    with open("./output/orders.csv", "wb") as code:
        code.write(request.content)

    # read the CSV-file
    out_list = []
    with open("./output/orders.csv") as file:
        rows = csv.reader(file)
        next(rows, None)
        for row in rows:
            out_list.append(row)
    return(out_list)


def open_the_robot_website(url):
    browser.open_available_browser(url)

def close_the_modal_window():
    browser.click_button_when_visible("css:.btn-dark")

def fill_the_form(data_row):
    browser.select_from_list_by_value("css:#head", data_row[1])
    browser.select_radio_button("body", data_row[2])
    browser.input_text("css:.form-control", data_row[3])
    browser.press_keys("css:.form-control", "ENTER")
    browser.input_text("css:#address", data_row[4])

def preview_the_robot():
    browser.click_button("css:#preview")
    browser.wait_until_element_is_visible("css:#robot-preview-image")

def submit_the_order():
    browser.click_button("css:#order")
    # Handling possible error while ordering (just repeat after wait)
    if browser.is_element_visible("css:.alert-danger"):
        time.sleep(5)
        preview_the_robot()
        submit_the_order()

def store_the_reciept_as_a_PDF(order_number):
    browser.wait_until_element_is_visible("css:#receipt")
    order_html = browser.find_element("css:#receipt").get_attribute('innerHTML')
    pdf.html_to_pdf(order_html, f"./output/orders/order_{order_number}.pdf")

def take_a_screenshot_of_the_robot(order_number):
    browser.capture_element_screenshot("css:#robot-preview-image", f"./output/orders/order_{order_number}.png")

def embed_the_screenshot_to_the_PDF(order_number):
    
    # # Perhaps more correct and flexible way to add image, but I couldn't resize the image :(
    # list_of_files = [
    #     "./output/orders/order.pdf",
    #     "./output/orders/order.png",
    #     ]
    # pdf.add_files_to_pdf(files=list_of_files, target_document="./output/orders/output.pdf")

    # Another way - works fine
    pdf.add_watermark_image_to_pdf(
        image_path=f"./output/orders/order_{order_number}.png",
        source_path=f"./output/orders/order_{order_number}.pdf",
        output_path=f"./output/orders/order_{order_number}_final.pdf",
    )

    # Delete the source files
    try:
        os.remove(f"./output/orders/order_{order_number}.png")
        os.remove(f"./output/orders/order_{order_number}.pdf")
    except:
        print("Failed to delete source files!")

def go_to_another_order():
    browser.click_button("css:#order-another")

def create_a_ZIP():
    zip = Archive()

    zip.archive_folder_with_zip(
        folder="./output/orders/",
        archive_name="./output/all_the_orders.zip",
    )

def delete_source_folder_with_orders():
    try:
        shutil.rmtree("./output/orders/")
    except:
        print("Failed to delete source folder with orders!")

# Define a main() function that calls the other functions in order:
def main():
    try:
        csv_link = input_dialog()  # "https://robotsparebinindustries.com/orders.csv"
        orders = get_orders_from_csv(csv_link)
        site_link = read_from_the_local_vault()
        open_the_robot_website(site_link)
        for order in orders:
            close_the_modal_window()
            fill_the_form(order)
            preview_the_robot()
            submit_the_order()
            store_the_reciept_as_a_PDF(order[0])
            take_a_screenshot_of_the_robot(order[0])
            embed_the_screenshot_to_the_PDF(order[0])
            go_to_another_order()
        create_a_ZIP()
        delete_source_folder_with_orders()
    except:
        print('Something was going wrong, the robot has been rejected :(')
    finally:
        browser.close_all_browsers()


# Call the main() function, checking that we are running as a stand-alone script:
if __name__ == "__main__":
    main()
