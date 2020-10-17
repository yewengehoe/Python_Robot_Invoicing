# set the webpage to half screen before running the robot
# https://next.waveapps.com/b3ec9290-2f2d-4e78-be62-00222ea08c1e/invoices
# see the screenshot file begin_page.png
import pandas as pd
import openpyxl
import input_data_function
import pyautogui
import time

df = pd.read_excel('multiple_items_jobs_to_bill.xlsx')

path = r"/home/yeh/PycharmProjects/myOperation/sorted_customer.xlsx"
writer = pd.ExcelWriter(path, engine='xlsxwriter')

customer_array_list = df.Customer.unique()

for customer in customer_array_list:
    customer_rows_extract = df.loc[(df.Customer == customer)]
    customer_rows_extract.to_excel(customer + '.xlsx', index=False)
    customer_rows_extract.to_excel(writer, sheet_name=customer, index=False)

writer.save()
writer.close()

time.sleep(1.0)
### Using newly created Excel file
for customer in customer_array_list:
    print("Generate invoice of ", customer)
    wb = openpyxl.load_workbook('sorted_customer.xlsx')
    # pick the sheet of current customer to operate
    sheet = wb[customer]
    # check number of rows/items
    row_count = sheet.max_row

    # Create a new invoice
    input_data_function.goto_add_invoice()

    i = 1
    PO = ""
    while i < row_count:
        # prepare the data to input
        PO = str(sheet['B' + str(i + 1)].value)
        print("PO number: ", PO)
        item = (sheet['C' + str(i + 1)].value)
        quantity = str((sheet['D' + str(i + 1)].value))
        price = str((sheet['E' + str(i + 1)].value))

        if i == 1:
            input_data_function.input_customer(customer)

        if PO != "None":
            pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/po_number.png', confidence=0.8))
            pyautogui.write(PO)

        input_data_function.inputItem(item, quantity, price)
        i = i + 1

    time.sleep(1.5)
    input_data_function.save_and_download_invoice()
    time.sleep(4.0)
    print("Invoice generated for ", customer)

# click to the page that shows the latest invoices
pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/main_menu.png', confidence=0.8))
pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/sidebar_invoices.png', confidence=0.8))
pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/close_main_menu.png', confidence=0.8))
pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/filter_from.png', confidence=0.8))
pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/filter_today_blue.png', confidence=0.8))
