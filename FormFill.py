
from SiteModule import SiteCommands
from selenium import webdriver
import time

'''
used to store general info about the current order.
non-xlsx variables store the reference to the site element, which is used
to send keys to for form filling. The xlsx variables are used for storing
the values from the spreadsheet
'''
class OrderInfo(object):
    def __init__(self): #excel values for current order(store info)
        self.store_id_xlsx = None
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

    def __str__(self):
        return '{}'.format(self.store_id)


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

        self.tmo_program_name = 'ssp16.8_tmo'
        self.partial_store_name = 'TMO-'
        self.full_store_name = 'T-Mobile Store '
        self.sheet_categories = {'store id' : 2,
                                 'address 1' : 30,
                                 'address 2' : 31,
                                 'zip' : 33}

        #the 'e_default' are values that are default in the form
        self.e_default_first_name = 'ATTN: Store'
        self.e_default_last_name = 'Operations Associate'
        self.e_default_email = 'aaa@viennachannels.com'
        self.e_default_phone = '409-622-3620' #public vienna number

        #keys used to navigate through submenus for the default form fill
        self.keys_default_ship_service = 'fff' #2 day shipping

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

    def get_form_elements(self, browser, form_elements):
        return None

    def fill_form(self, orders):
        return None



class FormFill(object):
    def __init__(self):
        self.username = 'zniemann'
        self.password = 'paluxy61'
        self.browser = webdriver.Firefox()
        self.program_form = None #form filler data; specified under 'navigate_to_form()'
        self.store_info = OrderInfo() #info for the current store
        self.store_row = None #row the store number is found on. All above data is pulled
                              #from this row

        self.wb = None #current workbook
        self.ws = None #current sheet in workbook
        self.all_stores = [] #list of all store numbers that have orders to be placed
                             #appended will be OrderInfo() objects

        self.supported_forms = [
            'tmo 16.8',
            ]


    #login to site where the form to fill exists
    def login(self, url, user_login, user_pass): #login first
        self.browser.get(url)
        username = SiteCommands.find_element_by_name(self.browser, 'username')
        password = SiteCommands.find_element_by_name(self.browser, 'password')

        SiteCommands.send_keys(username, user_login, 0) #enters username into field
        SiteCommands.send_keys(password, user_pass, 0) #enters password into field
        password.submit() #logs user in

        return None

    def navigate_to_form(self, url, form_name): #use after logged in
        SiteCommands.navigate_to_url(self.browser, url, 0.1)

        if form_name in self.supported_forms:
            if form_name == 'tmo 16.8':
                self.program_form = TmoSupport()
                self.program_form.navigate_to_form(self.browser)

        return None


    def get_form_elements(self, form_name): #use after navigating to form page
        if form_name in self.supported_forms:
            if form_name == 'tmo 16.8':
                pass #TmoSupport.get_form_elements(self.browser, )

        return None

    #adds all the stores to the store list
    def store_setup(self, stores_list):
        for store in stores_list:
            self.all_stores.append(store)

        return None
