# set the webpage to half screen before running the robot
# https://next.waveapps.com/b3ec9290-2f2d-4e78-be62-00222ea08c1e/invoices
# see the screenshot file begin_page.png
import pandas as pd
import openpyxl
import pyautogui
import time

pyautogui.PAUSE = 1.8
pyautogui.FAILSAFE = True

df = pd.read_excel('Jobs_to_bill.xlsx')
customer_array_list = df.Customer.unique()
print(customer_array_list)

time.sleep(2.0)
i = 1
for customer in customer_array_list:
    PO = ""
    print("Generate invoice of ", customer)
    wb = openpyxl.load_workbook('single_item_jobs_to_bill.xlsx')
    sheet = wb["Sheet1"]
    # time.sleep(1.0)
    # print(sheet)
    PO = str(sheet['B' + str(i + 1)].value)
    print("PO number: ", PO)
    item = (sheet['C' + str(i + 1)].value)
    quantity = str((sheet['D' + str(i + 1)].value))
    price = str((sheet['E' + str(i + 1)].value))

    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/main_menu.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/sidebar_invoices.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/close_main_menu.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/create_inv.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/add_a_customer.png', confidence=0.8))
    pyautogui.write(customer)
    pyautogui.press('enter')

    if PO != "None":
        pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/po_number.png', confidence=0.8))
        pyautogui.write(PO)

    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/add_an_item.png', confidence=0.8))
    pyautogui.write('a')
    pyautogui.press('enter')
    pyautogui.write(item)
    pyautogui.press('tab')
    pyautogui.write(quantity)
    pyautogui.press('tab')
    pyautogui.write(price)
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/save_and_cont.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/approve_draft.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/mark_as_sent.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/more_actions.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/export_as_PDF.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/download_PDF.png', confidence=0.8))
    i = i + 1
    time.sleep(4.0)
    print("Invoice generated for ", customer)
