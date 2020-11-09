import xlrd
import tkinter as tk
from tkinter import filedialog
from tkinter import *
import os
import sys
import time
import socket
import random #delays inevitability of captcha
from selenium import webdriver #Browser automation
from selenium.webdriver.support.ui import Select #selects items
from selenium.webdriver.chrome.options import Options #google chrome options
from selenium.common.exceptions import NoSuchElementException #for captcha
from xlwt import Workbook
from xlutils.copy import copy
import threading


options = Options()
options.add_argument("--log-level=3")  #disables/suppresses chrome error warnings for selenium
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.headless = True

driver_path = "chromedriver_win32.exe"

#C:\\Users\\PA40008372\\PycharmProjects\\pythonProject\\venv\\Lib\\site-packages\\chromedriver_py\\

root = Tk()
root.geometry('800x200')

file_Path = "FILE NOT SELECTED! PLEASE SELECT ONLY .XLS Files"

def file_opener():
    global file_Path
    file_Path = filedialog.askopenfilename()
    T.delete(tk.CURRENT,tk.END)
    T.insert(tk.CURRENT, file_Path)

def Check():
    root.withdraw()
    root1 = Tk()
    def stop_fun():
        root1.withdraw()
        quit()
        sys.exit()

    root1.geometry('400x200')
    M = Button(root1, text='EXIT', command=lambda: stop_fun())
    R = Text(root1, height = 5, width = 60)
    M.pack()
    R.pack()
    wb = xlrd.open_workbook(file_Path)
    sheet = wb.sheet_by_index(0)
    count = 0
    wb1 = copy(wb)
    s = wb1.get_sheet(0)
    total = sheet.nrows

    for i in range(1, total):
        flag = 0
        categorisation = ""
        website = sheet.cell_value(i, 0)
        try:
            if website != "":
                driver = webdriver.Chrome(driver_path, options=options)
                driver.get('https://sitereview.bluecoat.com/#/''https://sitereview.bluecoat.com/#/')
                #driver.get('https://sitereview.bluecoat.com/#/captcha')
                inputElement = driver.find_element_by_id("txtUrl")
                inputElement.send_keys(website)
                inputElement.submit()
                time.sleep(2)
                categorisation = driver.find_element_by_class_name('clickable-category').text
                time.sleep(2)
                driver.close()
            else:
                categorisation = "No Information"
        except NoSuchElementException as e:
            try:
                categorisation = "Error"
                if (bool(driver.find_element_by_id('imgCaptcha')) == True):
                    #print("Reason: CAPTCHA appeared! Please fill in and try again later!")
                    R.insert(tk.CURRENT, "Error: CAPTCHA appeared! Please fill in and try again later!")
                    flag  = 1
                    driver.close()
                    break
                    #mainloop(1)
                    time.sleep(3)
            except NoSuchElementException:
                categorisation = "Error"
                driver.close()
        if categorisation == "":
            time.sleep(1)
            #print("Entered sleeping")
        else:
            #print(categorisation)
            s.write(i, 2, categorisation)
            wb1.save(file_Path)
        time.sleep(8)
    if flag == 0:
            R.insert(tk.CURRENT, "Completed!")
    mainloop(1)


    #print("Done")


x = Button(root, text ='Select a file', command = lambda:file_opener())
y = Button(root, text ='Run', command = lambda:Check())

T = Text(root, height = 5, width = 92)
T.insert(tk.END, file_Path)

x.pack()
y.pack()
T.pack()
mainloop()