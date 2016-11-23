from selenium import webdriver
from ExcelModule import ExcelCommands
import openpyxl
import time
import datetime
from SiteModule import SiteCommands
import FormFill
from FormFill import FormFill
import sys

def get_stores():
    if len(sys.argv) > 1:
        return sys.argv[1:]

    return None

def run_form_fill(username, password, login_url, program_url, xlsx_name):
    form_data = FormFill()
    form_data.all_stores = get_stores()

    if form_data.all_stores is None:
        form_data.all_stores = example_stores

    form_data.get_form_program('tmo 16.8')

    form_data.load_workbook(xlsx_file) #loads excel file
    form_data.get_excel_store_info()

    form_data.login(login_url, username, password) #logs user into site where form is located
    form_data.navigate_to_form(program_url, 'tmo 16.8') #navigates to form
    form_data.get_form_elements() #stores site elements from form
    form_data.write_to_form()

def main():
    example_stores = ['4176', '4200', '9513']

    username = 'zniemann'
    password = 'paluxy61'
    login_url = r'https://www.viennachannels.com/adminlogin.php'
    program_url = r'https://www.viennachannels.com/vca/user_mgmt.php'
    xlsx_file = r'tmo_shipment_details.xlsx'

    form_data = FormFill()
    form_data.all_stores = get_stores()

    if form_data.all_stores is None:
        form_data.all_stores = example_stores

    form_data.get_form_program('tmo 16.8')

    form_data.load_workbook(xlsx_file) #loads excel file
    form_data.get_excel_store_info()

    form_data.login(login_url, username, password) #logs user into site where form is located
    form_data.navigate_to_form(program_url, 'tmo 16.8') #navigates to form
    form_data.get_form_elements() #stores site elements from form
    form_data.write_to_form()

    form_data.browser.close()
    return None

main()
