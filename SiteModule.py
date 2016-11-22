import time
import datetime

date = datetime.datetime.now()
time_out_limit = 10

class SiteCommands(object):
    '''
    def __init__(self):
        self.store_id = None
        self.order_type = None
        self.company_name = None
        self.first_name = None
        self.last_name = None
        self.address_1 = None
        self.address_2 = None
        self.city = None
        self.state = None
        self.zip = None
        self.email = None
        self.confirm_email = None
        self.phone = None
        self.ship_service = None
        '''


    def find_element_by_link_text_click(browser, txt):
        time_out = (date.now().second + time_out_limit) % 60

        while True:
            try:
                if date.now().second == time_out:
                    print('Error: Site timed out')
                    break

                browser.find_element_by_link_text(txt).click()
                break

            except Exception as e:
                pass

        return None

    def find_element_by_id_click(browser, id_name):
        time_out = (date.now().second + time_out_limit) % 60

        while True:
            try:
                if date.now().second == time_out:
                    print('Error: Site timed out')
                    break

                browser.find_element_by_id(id_name).click()
                break

            except Exception as e:
                continue

        return None

    def find_element_by_name(browser, name):
        found_element = None
        time_out = (date.now().second + time_out_limit) % 60

        while True:
            try:
                if date.now().second == time_out:
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

    def cycle_element_options(element, options, pause_time):

        for option in options:
            time.sleep(pause_time) #pauses for set amount of time before sending key
            element.send_keys(option)

        return None

    def send_keys(element, keys, pause_time):
        for key in keys: #goes through all keys to send, one at a time
            if pause_time > 0:
                time.sleep(pause_time)

            element.send_keys(key)

        return None

    def navigate_to_url(browser, url, pause_time):
        time_out = (date.now().second + time_out_limit) % 60

        while True:
            try:
                if date.now().second == time_out:
                    print('Error: Site timed out')
                    break

                browser.get(url)
                time.sleep(pause_time)
                break

            except Exception as e:
                pass

        return None
