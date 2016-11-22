from selenium import webdriver
from ExcelModule import ExcelCommands
import openpyxl
import time
import datetime
from SiteModule import SiteCommands
import FormFill
from FormFill import FormFill

date_obj = datetime.datetime.now()
time_out_limit = 10 #number of seconds before the site times out from trying to find an element

def login(browser, url, user_login, user_pass):
    browser.get(url)
    username = find_element_by_name(browser, 'username')
    password = find_element_by_name(browser, 'password')

    username.send_keys(user_login)
    password.send_keys(user_pass)
    password.submit()

    return None

def find_element_by_link_text_click(browser, txt):
    time_out = (date_obj.now().second + time_out_limit) % 60
    error_info = '' #sends error info through return

    while True:
        try:
            if date_obj.now().second == time_out:
                print('Error: Site timed out')
                break

            browser.find_element_by_link_text(txt).click()
            break

        except Exception as e:
            pass

    return error_info

def find_element_by_id_click(browser, id_name):
    time_out = (date_obj.now().second + time_out_limit) % 60
    error_info = '' #sends error info through return

    while True:
        try:
            if date_obj.now().second == time_out:
                print('Error: Site timed out')
                break

            browser.find_element_by_id(id_name).click()
            break

        except Exception as e:
            pass

    return error_info

def find_element_by_name(browser, name):
    found_element = None
    time_out = (date_obj.now().second + time_out_limit) % 60

    while True:
        try:
            if date_obj.now().second == time_out:
                print('Error: Site timed out')
                break

            found_element = browser.find_element_by_name(name)
            break

        except Exception as e:
            pass

    return found_element

def find_element_by_id(browser, id):
    found_element = None
    time_out = (date_obj.now().second + time_out_limit) % 60

    while True:
        try:
            if date_obj.now().second == time_out:
                print('Error: Site timed out')
                break

            found_element = browser.find_element_by_id(id)
            break

        except Exception as e:
            pass

    return found_element

def program_selection(browser, url):
    browser.get(url)
    time.sleep(1)
    program_name = find_element_by_name(browser, 'name') #browser.find_element_by_name('name')
    program_name.send_keys('ssp16.8_tmo')
    program_name.submit()
    time.sleep(1.5)

    find_element_by_id_click(browser, 'userrow-22661')
    find_element_by_id_click(browser, 'umlogin')

    print('found form click elements')
    time.sleep(1.5)
    find_element_by_link_text_click(browser, 'click here')

    print('past link text click')

    return None

def cycle_element_options(element, options, pause_time):

    for option in options:
        time.sleep(pause_time) #pauses for set amount of time before sending key
        element.send_keys(option)

    return None

def find_site_elements(browser):
    if browser is not None:
        print('start collecting browser elements')
        #store_id = browser.find_element_by_id('storeid')
        store_id = find_element_by_id(browser, 'storeid')
        #order_type = browser.find_element_by_id('sgevordertypeid')
        order_type = find_element_by_id(browser, 'sgevordertypeid')

        company_name = find_element_by_id(browser, 'eucompany') #browser.find_element_by_id('eucompany')

        first_name = find_element_by_id(browser, 'eufname') #browser.find_element_by_id('eufname')
        last_name = find_element_by_id(browser, 'eulname') #browser.find_element_by_id('eulname')

        address_1 = find_element_by_id(browser, 'euaddr1') #browser.find_element_by_id('euaddr1')
        address_2 = find_element_by_id(browser, 'euaddr2') #browser.find_element_by_id('euaddr2')

        city = find_element_by_id(browser, 'eucity') #browser.find_element_by_id('eucity')
        #state = browser.find_element_by_id('')
        zipcode = find_element_by_id(browser, 'euzip') #browser.find_element_by_id('euzip')
        email_address = find_element_by_id(browser, 'euemail1' ) #browser.find_element_by_id('euemail1')
        confirm_email = find_element_by_id(browser, 'euemail2') #browser.find_element_by_id('euemail2')
        phone = find_element_by_id(browser, 'euphone') #browser.find_element_by_id('euphone')
        ship_service = find_element_by_id(browser, 'svcidus') #browser.find_element_by_id('svcidus')

        #determine this by googling the address, finding what type of property the
        #recepient is on. If unable to determine, send a dialog message to the user
        #residential = browser.find_element_by_id('')

        #ship_service = browser.find_element_by_id('')
        #return_label = browser.find_element_by_id('')

        #it is intentional that there is no submit variable.
        print('loaded browser elements')

        return None


def tmo_form_fill(browser = None, wb_name = ''):
    try:
        print('entered: tmo_form_fill()')
        #browser.get(url)
        wb = openpyxl.load_workbook(wb_name)
        print('load workbook')
        ws = wb.active #active sheet in workbook
        store_num = '319'
        store_to_find = 'TMO-' + store_num
        store_full_name = 'T-Mobile Store '

        store_id_xlsx = None
        store_row = 0



        sheet_categories = {'store id' : 2,
                            'address 1' : 30,
                            'address 2' : 31,
                            'zip' : 33}

        print('found store id element on the form')

        store_id_xlsx = ExcelCommands.is_value_in_sheet(wb.active,
                                                 [2, 2],
                                                 [1, -1], store_to_find)
        store_row = store_id_xlsx[0]

        address_1_xlsx = ExcelCommands.get_cell_value(wb.active, sheet_categories['address 1'], store_row)
        city_xlsx = ExcelCommands.get_cell_value(wb.active, sheet_categories['address 2'], store_row)
        zip_xlsx = ExcelCommands.get_cell_value(wb.active, sheet_categories['zip'], store_row)
        default_first_name = 'ATTN: Store'
        default_last_name = 'Operations Associate'
        default_email = 'aaa@viennachannels.com'
        default_phone = '409-622-3620'
        default_ship_service = 'fff' #keys to send for 2 day shipping

        #make so it also keeps the value of the cell, and not just the position
        #of the cell
        #address_1_xlsx = ExcelCommands.is_value_in_sheet

        print('Excel sheet with the store id has been loaded')
        print('store id: {}'.format(store_id_xlsx) )

            #read in all other values based on this row, and the column category

        if browser is not None:
            print('start collecting browser elements')
            #store_id = browser.find_element_by_id('storeid')
            store_id = find_element_by_id(browser, 'storeid')
            #order_type = browser.find_element_by_id('sgevordertypeid')
            order_type = find_element_by_id(browser, 'sgevordertypeid')

            company_name = find_element_by_id(browser, 'eucompany') #browser.find_element_by_id('eucompany')

            first_name = find_element_by_id(browser, 'eufname') #browser.find_element_by_id('eufname')
            last_name = find_element_by_id(browser, 'eulname') #browser.find_element_by_id('eulname')

            address_1 = find_element_by_id(browser, 'euaddr1') #browser.find_element_by_id('euaddr1')
            address_2 = find_element_by_id(browser, 'euaddr2') #browser.find_element_by_id('euaddr2')

            city = find_element_by_id(browser, 'eucity') #browser.find_element_by_id('eucity')
            #state = browser.find_element_by_id('')
            zipcode = find_element_by_id(browser, 'euzip') #browser.find_element_by_id('euzip')
            email_address = find_element_by_id(browser, 'euemail1' ) #browser.find_element_by_id('euemail1')
            confirm_email = find_element_by_id(browser, 'euemail2') #browser.find_element_by_id('euemail2')
            phone = find_element_by_id(browser, 'euphone') #browser.find_element_by_id('euphone')
            ship_service = find_element_by_id(browser, 'svcidus') #browser.find_element_by_id('svcidus')

            #determine this by googling the address, finding what type of property the
            #recepient is on. If unable to determine, send a dialog message to the user
            #residential = browser.find_element_by_id('')

            #ship_service = browser.find_element_by_id('')
            #return_label = browser.find_element_by_id('')

            #it is intentional that there is no submit variable.
            print('loaded browser elements')

        #read in above values from the excel sheet
        if store_id_xlsx is not None:
            print('store id in excel is not none')
            store_id.send_keys(store_to_find)
            company_name.send_keys(store_full_name + store_num)
            first_name.send_keys(default_first_name)
            last_name.send_keys(default_last_name)
            address_1.send_keys( str(address_1_xlsx) )
            city.send_keys( str(city_xlsx) )
            zipcode.send_keys( str(zip_xlsx) )
            email_address.send_keys( str(default_email) )
            confirm_email.send_keys( str(default_email) )
            phone.send_keys( str(default_phone) )

            cycle_element_options(order_type, ['r'], 0.1)
            cycle_element_options(ship_service, default_ship_service, 0.1)

            print('finished sending excel keys')


    except Exception as e:
        print('tmo form fill error')
        print('{}'.format(e.with_traceback) )

    return None

def main():
    #browser = webdriver.Firefox()
    username = 'zniemann'
    password = 'paluxy61'
    login_url = r'https://www.viennachannels.com/adminlogin.php'
    program_url = r'https://www.viennachannels.com/vca/user_mgmt.php'
    xlsx_file = r'tmo_shipment_details.xlsx'

    form_data = FormFill()
    form_data.load_workbook(xlsx_file) #loads excel file
    
    form_data.login(login_url, username, password) #logs user into site where form is located
    form_data.navigate_to_form(program_url, 'tmo 16.8') #navigates to form
    form_data.get_form_elements() #stores site elements from form
    #login(browser, login_url, username, password)
    #program_selection(browser, program_url)
    #tmo_form_fill(browser, xlsx_file)
    return None

main()
