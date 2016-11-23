
from SiteModule import SiteCommands
from ExcelModule import ExcelCommands
from selenium import webdriver
import time
import openpyxl

'''
used to store general info about the current order.
non-xlsx variables store the reference to the site element, which is used
to send keys to for form filling. The xlsx variables are used for storing
the values from the spreadsheet
'''
class OrderInfo(object):
    def __init__(self): #excel values for current order(store info)
        self.store_row = None #row the store number is found on. All above data is pulled
                              #from this row

        self.store_num = None #ex. '952' from 'TMO-952'
        self.store_id_xlsx = None #ex 'TMO-952'
        self.order_type_xlsx = None
        self.company_name_xlsx = None
        self.first_name_xlsx = None
        self.last_name_xlsx = None
        self.address_1_xlsx = None
        self.address_2_xlsx = None
        self.city_xlsx = None
        self.state_xlsx = None
        self.zip_xlsx = None
        self.email_xlsx = None
        self.phone_xlsx = None
        self.ship_service_xlsx = None

    def __str__(self):
        return '{}'.format(self.store_id_xlsx)


'''program specific classes are used on a case by case basis. They hold specific information
to the class, and do not need to be instantiated.
'''
class TmoSupport(object):
    def __init__(self):
        #each 'e_' value holds the site's corresponding field, used to
        #click and send keys to the fields
        self.e_store_id = None
        self.e_order_type = None
        self.e_company_name = None
        self.e_first_name = None
        self.e_last_name = None
        self.e_address_1 = None
        self.e_address_2 = None
        self.e_city = None
        self.e_state = None
        self.e_zip = None
        self.e_email = None
        self.e_confirm_email = None
        self.e_phone = None
        self.e_ship_service = None

        self.program_url = r'https://www.viennachannels.com/vca/user_mgmt.php'

        self.tmo_program_name = 'ssp16.8_tmo'
        self.partial_store_name = 'TMO-'
        self.full_store_name = 'T-Mobile Store '
        self.sheet_categories = {'store id' : 2,
                                 'address 1' : 30,
                                 'city' : 31,
                                 'state' : 32,
                                 'zip' : 33}

        #the 'e_default' are values that are default in the form
        self.e_default_first_name = 'ATTN: Store'
        self.e_default_last_name = 'Operations Associate'
        self.e_default_email = 'aaa@viennachannels.com'
        self.e_default_phone = '409-622-3620' #public vienna number

        #keys used to navigate through submenus for the default form fill
        self.keys_default_ship_service = 'fff' #2 day shipping
        self.keys_default_order_type = 'r' #replacement order type

    #navigates to the form for t-mobile support
    def navigate_to_form(self, browser): #needs to support all tmo programs, not just 16.8
        program_name = SiteCommands.find_element_by_name(browser, 'name') #browser.find_element_by_name('name')
        SiteCommands.send_keys(program_name, 'ssp16.8_tmo', 0)
        program_name.submit()
        time.sleep(1.5)

        SiteCommands.find_element_by_id_click(browser, 'userrow-22661')
        SiteCommands.find_element_by_id_click(browser, 'umlogin')

        print('found form click elements')
        time.sleep(1.5)
        SiteCommands.find_element_by_link_text_click(browser, 'click here')
        return None

    def get_form_elements(self, browser):
        self.e_store_id = SiteCommands.find_element_by_id(browser, 'storeid')
        self.e_order_type = SiteCommands.find_element_by_id(browser, 'sgevordertypeid')

        self.e_company_name = SiteCommands.find_element_by_id(browser, 'eucompany') #browser.find_element_by_id('eucompany')

        self.e_first_name = SiteCommands.find_element_by_id(browser, 'eufname') #browser.find_element_by_id('eufname')
        self.e_last_name = SiteCommands.find_element_by_id(browser, 'eulname') #browser.find_element_by_id('eulname')

        self.e_address_1 = SiteCommands.find_element_by_id(browser, 'euaddr1') #browser.find_element_by_id('euaddr1')
        self.e_address_2 = SiteCommands.find_element_by_id(browser, 'euaddr2') #browser.find_element_by_id('euaddr2')

        self.e_city = SiteCommands.find_element_by_id(browser, 'eucity') #browser.find_element_by_id('eucity')
        #state = browser.find_element_by_id('')
        self.e_zip = SiteCommands.find_element_by_id(browser, 'euzip') #browser.find_element_by_id('euzip')
        self.e_email = SiteCommands.find_element_by_id(browser, 'euemail1' ) #browser.find_element_by_id('euemail1')
        self.e_confirm_email = SiteCommands.find_element_by_id(browser, 'euemail2') #browser.find_element_by_id('euemail2')
        self.e_phone = SiteCommands.find_element_by_id(browser, 'euphone') #browser.find_element_by_id('euphone')
        self.e_ship_service = SiteCommands.find_element_by_id(browser, 'svcidus') #browser.find_element_by_id('svcidus')
        return None

#main class. This is what is instantiated by you to use for the form filler.
class FormFill(object):
    #optional
    def load_page(self, pause_time):
        self.browser.refresh()
        time.sleep(pause_time)

        return None

    #FormFill process in top down order
    def __init__(self):
        self.username = 'zniemann'
        self.password = 'paluxy61'
        self.browser = webdriver.Firefox()
        self.program_form = None #form filler data; specified under 'navigate_to_form()'
        #self.store_info = None #info for the current store

        self.wb = None #current workbook
        self.ws = None #current sheet in workbook
        self.ws_total_rows = 0 #total number of rows in active ws
        self.ws_total_cols = 0 #total number of columns in active ws

        self.all_stores = [] #list of all store numbers that have orders to be placed
                             #appended will be OrderInfo() objects

        self.supported_forms = [
            'tmo 16.8',
            ]

    def load_workbook(self, wb_name): #do at the beginning of program
        self.wb = openpyxl.load_workbook(wb_name)
        self.ws = self.wb.active
        self.ws_total_rows = len( list(self.ws.rows) ) #total rows in the active worksheet
        self.ws_total_cols = len( list(self.ws.columns) ) #total columns in the active worksheet

        return None

    def get_form_program(self, program_name): #gets the current program
        if program_name in self.supported_forms:
            if program_name == 'tmo 16.8':
                self.program_form = TmoSupport()
                print('current program: Tmo support')

        return None

    def get_excel_store_info(self): #reads store address from spreadsheet and stores for later user_pass
        store_data = [] #holds an OrderInfo store data object for each store
        for i in range( len(self.all_stores) ):
            #self.all_stores[i] = self.program_form.partial_store_name + self.all_stores[i]

            #print('current store: {}'.format(self.all_stores[i]) )
            store_info = OrderInfo()
            store_info.store_num = self.all_stores[i]
            store_info.store_id_xlsx = self.program_form.partial_store_name + store_info.store_num

            temp = ExcelCommands.is_value_in_sheet(self.ws,
                                                   [self.program_form.sheet_categories['store id'],
                                                    self.program_form.sheet_categories['store id']],
                                                   [1, -1],
                                                    store_info.store_id_xlsx)

            if temp is not None: #get info for store
                store_info.store_num = self.all_stores[i] #store number
                store_info.store_row = temp[0] #row of current store info

                store_info.store_id_xlsx = ExcelCommands.get_cell_value(self.ws,
                                                                             self.program_form.sheet_categories['store id'],
                                                                             store_info.store_row)

                store_info.address_1_xlsx = ExcelCommands.get_cell_value(self.ws,
                                                                              self.program_form.sheet_categories['address 1'],
                                                                              store_info.store_row)

                store_info.city_xlsx = ExcelCommands.get_cell_value(self.ws,
                                                                         self.program_form.sheet_categories['city'],
                                                                         store_info.store_row)

                store_info.state_xlsx = ExcelCommands.get_cell_value(self.ws,
                                                                          self.program_form.sheet_categories['state'],
                                                                          store_info.store_row)

                store_info.zip_xlsx = ExcelCommands.get_cell_value(self.ws,
                                                                        self.program_form.sheet_categories['zip'],
                                                                        store_info.store_row)

                store_info.order_type_xlsx = self.program_form.keys_default_order_type
                store_info.company_name_xlsx = self.program_form.full_store_name + store_info.store_num
                store_info.first_name_xlsx = self.program_form.e_default_first_name
                store_info.last_name_xlsx = self.program_form.e_default_last_name
                store_info.email_xlsx = self.program_form.e_default_email
                store_info.phone_xlsx = self.program_form.e_default_phone
                store_info.ship_service_xlsx = self.program_form.keys_default_ship_service

                #store_data.append(self.store_info)
                self.all_stores[i] = store_info

            else:
                print('FormFill() --> get_excel_store_info(): temp is None')

        return None

    #login to site where the form to fill exists
    def login(self, url, user_login, user_pass): #login first
        self.browser.get(url)
        username = SiteCommands.find_element_by_name(self.browser, 'username')
        password = SiteCommands.find_element_by_name(self.browser, 'password')

        pause_time = 0 #time to pause between sending each key
        SiteCommands.send_keys(username, user_login, pause_time) #enters username into field
        SiteCommands.send_keys(password, user_pass, pause_time) #enters password into field
        password.submit() #logs user in

        return None

    def navigate_to_form(self, url, form_name): #use after logged in
        SiteCommands.navigate_to_url(self.browser, url, 0.1)
        self.program_form.navigate_to_form(self.browser)

        return None

    def get_form_elements(self): #use after navigating to form page
        self.program_form.get_form_elements(self.browser)
        #^THIS CALL AND PARAMETERS SHOULD BE THE SAME, REGARDLESS OF THE PROGRAM.
        #THIS WILL KEEP IT SIMPLE SO THERE WON'T HAVE TO BE A TON OF CHECKS TO DETERMINE WHICH FORM VALUES TO USE.

        return None

    def write_to_form(self): #writes all info from spreadsheet & order to form
        for i in range( len(self.all_stores) ):

            print('store id')
            SiteCommands.send_keys(self.program_form.e_store_id,
                                   self.all_stores[i].store_id_xlsx,
                                   0)

            print('order type')
            SiteCommands.send_keys(self.program_form.e_order_type,
                                   self.all_stores[i].order_type_xlsx,
                                   0.2)

            print('company name')
            SiteCommands.send_keys(self.program_form.e_company_name,
                                   self.all_stores[i].company_name_xlsx,
                                   0)

            print('first name: {}'.format(self.all_stores[i].first_name_xlsx) )
            SiteCommands.send_keys(self.program_form.e_first_name,
                                   self.all_stores[i].first_name_xlsx,
                                   0)

            print('last name')
            SiteCommands.send_keys(self.program_form.e_last_name,
                                   self.all_stores[i].last_name_xlsx,
                                   0)

            print('address 1')
            SiteCommands.send_keys(self.program_form.e_address_1,
                                   self.all_stores[i].address_1_xlsx,
                                   0)

            print('city')
            SiteCommands.send_keys(self.program_form.e_city,
                                   self.all_stores[i].city_xlsx,
                                   0)

            #state

            print('zip')
            SiteCommands.send_keys(self.program_form.e_zip,
                                   self.all_stores[i].zip_xlsx,
                                   0)

            print('email')
            SiteCommands.send_keys(self.program_form.e_email,
                                   self.all_stores[i].email_xlsx,
                                   0)

            print('confirm email')
            SiteCommands.send_keys(self.program_form.e_confirm_email,
                                   self.all_stores[i].email_xlsx,
                                   0)

            print('phone')
            SiteCommands.send_keys(self.program_form.e_phone,
                                   self.all_stores[i].phone_xlsx,
                                   0)

            print('ship service')
            SiteCommands.send_keys(self.program_form.e_ship_service,
                                   self.all_stores[i].ship_service_xlsx,
                                   0.2)

            input() #pause
            self.load_page(2.0)

        return None
