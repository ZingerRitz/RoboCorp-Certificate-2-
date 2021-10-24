#certification Level 2 
import os
import selenium
import tkinter


from Browser import Browser
from Browser.utils.data_types import SelectAttribute
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Archive import Archive
from selenium.webdriver.support.ui import Select
from RPA.Browser.Selenium import Selenium
from RPA.Tables import Tables
from tkinter import *
from tkinter import messagebox
from RPA.Robocorp.Vault import Vault

Surf = Browser()


# Open Browser
def Open_browser():
    vault = Vault()
    small = vault.get_secret("website")
    tryonce = small["url"]
    Surf.open_browser(tryonce)


# Login with vault feature
def log_in():
    Surf.type_text("css=#username", "maria")
    Surf.type_secret("css=#password", "thoushallnotpass")
    Surf.click("text=Log in")


# download the CSV File
def download_the_excel_file():
    http = HTTP()
    http.download(
        url="https://robotsparebinindustries.com/orders.csv",
        overwrite=True)


# Move to an other tab
def goto_otherTab():
    Surf.click("text=Order your robot!")
    Surf.wait_for_elements_state("css=.modal-header")
    Surf.click("text=OK")


# read CSV and populate all the values into respective fields
def readCsv():
    csv1 = Tables()
    orderInfos = csv1.read_table_from_csv("orders.csv",
                                          columns=["Order number", "Head",
                                                   "Body", "Legs", "Address"])
    for orderList in orderInfos:
        selectRobot(orderList)
        submit(orderList)
        takeScreenShot()
        createInvoice()
        OrderAnother()
        addScreenshot_toPdf(orderList)
        moveFilesIntoZip()


# filling all forms
def selectRobot(orderList):
    Surf.select_options_by("css=#head", SelectAttribute["value"],
                           str(orderList["Head"]))
    q=str(orderList["Body"])
    Surf.check_checkbox("css=#id-body-"+q)
    Surf.type_text("xpath=//html/body/div/div/div[1]/div/div[1]/form/div[3]/input",
                   str(orderList["Legs"]))
    Surf.type_text("css=#address", orderList["Address"])
    Surf.click("css=#preview")
    Surf.wait_for_elements_state("css=#robot-preview")


# Submit the order and Error handling
def submit(orderList):
    try:
        Surf.click("css=#order")
        Surf.wait_for_elements_state("xpath=//html/body/div/div/div[1]/div/div[1]/div/button")
    except Exception as e:
        Surf.wait_for_elements_state("xpath=//html/body/div/div/div[1]/div/div[1]/div")
        for i in range(0, 20):
            try:
                Surf.click("css=#order")
                Surf.wait_for_elements_state("xpath=//html/body/div/div/div[1]/div/div[1]/div/button")
            except Exception as e:
                pass
            else:
                break


# take a screenshot
def takeScreenShot():
    Surf.take_screenshot(
            filename=f"{os.getcwd()}/order.png",
            selector="xpath=//html/body/div/div/div[1]/div/div[2]/div/div")


#append HTML data
def createInvoice():
    Surf.wait_for_elements_state("css=#order-another")
    invoice_start = Surf.get_property(
        selector="xpath=//html/body/div/div/div[1]/div/div[1]/div/div",
        property="outerHTML")
    pdf = PDF()
    pdf.html_to_pdf(invoice_start, "invoicerp.pdf")


# Place an other order 
def OrderAnother():
    Surf.click("xpath=//html/body/div/div/div[1]/div/div[1]/div/button")
    Surf.wait_for_elements_state("css=.modal-header")
    Surf.click("text=OK")


# Append HTML data with image into a PDF
def addScreenshot_toPdf(orderList):
    pd = PDF()
    activeList = ['invoicerp.pdf', 'order.png:x=0,y=0']
    pd.add_files_to_pdf(
        activeList,
        target_document="pngOutput/valid"+str(orderList["Order number"])+".pdf"
    )


# Move file into Zip
def moveFilesIntoZip():
    arc = Archive()
    arc.archive_folder_with_zip("pngOutput", "Output/zipped.zip", recursive=True)


def Logout():
    Surf.click("xpath=//html/body/div/header/div/div/span[2]/button")


def Close_B():
    Surf.close_browser()


# Message box with an option to logout or close the browser
# NOTE: if the Message box cant be seen on top, it would be behind the browser window
def dialup():
    wind = Tk()
    wind.eval('tk::PlaceWindow %s center' % wind.winfo_toplevel())
    wind.withdraw()

    if messagebox.askyesno('Question', 'Do you want to close the Browser') == True:
        Close_B()
    else:
        Logout()

    wind.deiconify()
    wind.destroy()
    wind.quit()


def main():
    Open_browser()
    log_in()
    download_the_excel_file()
    goto_otherTab()
    readCsv()
    dialup()


if __name__ == "__main__":
    main()
