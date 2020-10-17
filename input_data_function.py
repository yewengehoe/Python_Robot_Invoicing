import pyautogui
import time

pyautogui.PAUSE = 2.0
pyautogui.FAILSAFE = True

def goto_add_invoice():
    # time.sleep(0.5)
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/main_menu.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/sidebar_invoices.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/close_main_menu.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/create_inv.png', confidence=0.8))

def input_customer(customer_name):
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/add_a_customer.png', confidence=0.8))
    pyautogui.write(customer_name)
    pyautogui.press('enter')

def inputItem(item, quantity, price):
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/add_an_item.png', confidence=0.8))
    pyautogui.write('a')
    pyautogui.press('enter')
    pyautogui.write(item)
    pyautogui.press('tab')
    pyautogui.write(quantity)
    pyautogui.press('tab')
    pyautogui.write(price)

def save_and_download_invoice():
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/save_and_cont.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/approve_draft.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/mark_as_sent.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/more_actions.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/export_as_PDF.png', confidence=0.8))
    pyautogui.click(pyautogui.locateCenterOnScreen('./myInvoiceButton/download_PDF.png', confidence=0.8))
