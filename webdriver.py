''' This is the main method of interacting with SmartEtailin in the same way a human would. This should only be used in situations where a documented API does not exist... which is most situations.

start - creates a webdriver instance in headless mode. Can be specified as non-headless, specific window size, and to use a proxy

login - login to SE using credentials provided in config.yaml This ia a required step before most other fuctions

new_orders - opens orders page and checks all the boxes on the first page where the order status is "New Order Pending". This stages these orders to populate the official XML order export API. NOTE: unlike the official export method, this ignores every aspect other that the order status, so it will include orders prior to requesting fulfillment.

capture_payments - opens an order in SE and clicks capture payment for those that are made via Authorize.net

request_fulfillment - requests fulfillment for items on an order for QBP and Trek. HLC could be built in later, but It may be better to use their official API unless you really want feedback from SE's integration. HLC's API is well documented and reliable.

change_shipping - simple function to change the shipping type from Home Delivery to Ground Shipping to allow for QBP and Trek fulfillment of these orders. If you submit a fulfillment request when home delivery is set, SE will get a notification from Q that they refused the order.

get_discount_ID - pulls the discount description from an order's page

update_shipping_cost - In Development - method to change the "Additional Shipping" field of an item. This can be used to set up additional logic to your shipping policies. By default, SE allows dimensional weight, top level categories, and price thresholds. So if you want different shipping costs for different prices or types of bikes, you have to do it this way.

create_unique_discount - creates a unique discount code that is able to be used once on any item or order.

delete_old_discounts - removes "EXPIRED" discounts from commerce-discounts page. When using many unique discounts, that page can get unweily very fast. Currently, this is set to start at page 11 and run until no more are left.

get_product_map - pulls SKU/UPC/MPN product mapping file. This can be usefull for identifying which products are in your catalog and mapping correctly, or can be used to get itemIDs for modifying their pages.

update_item_notes - In Development - updates the notes of an item within an order.

update_ordernotes - Appends additional information to the bottom of the order notes.

get_config - parses config.yaml and returns a dictionary.
'''

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
import time
import traceback
import yaml
import re

def start(headless=True,size="1266,1382",proxy=False):
    options = Options()
    options.add_argument("--window-size="+size)
    if proxy == True:
        options.add_argument('--proxy-server=http://x.botproxy.net:8080')
        options.add_argument('--host-rules=MAP www.paypal.com 127.0.0.1')
        prefs = {"profile.managed_default_content_settings.images": 2}
        options.add_experimental_option("prefs", prefs)
    options.headless = headless
    driver = webdriver.Chrome('./chromedriver',chrome_options=options)
    return [0,driver]

def login(driver,un,pw,domain):
    try:
        url = 'https://www.'+domain+'/admin/login.cfm?CFTOKEN=0&cookie=yes'
        driver.get(url)
        print(driver.current_url)
        login = driver.find_element_by_name('Login')
        login.send_keys(un)
        login = driver.find_element_by_name('Pass')
        login.send_keys(pw)
        login.submit()
        if driver.current_url == 'https://www.'+domain+'/admin/index.cfm':
            print('SE Login Successful')
            return [0,driver,driver.current_url]
        else:
            print('Login Error, resulted in page: {}'.format(driver.current_url))
            return [2,driver,driver.current_url]
    except: 
        print('Script Error')
        return [1,driver,'']

def new_orders(driver,domain):
    try:
        url = 'https://www.'+domain+'/admin/orders/openorders.cfm?Status=1&startYear=1900&sort=Created&ordertype=desc'
        driver.get(url)
        count = 0
        try:
            for i in range(3,104):
                status = Select(driver.find_element_by_xpath('//*[@id="main-container"]/form/div/div/div[1]/table/tbody/tr[{}]/td[9]/select'.format(i))).first_selected_option.text
                if status == "New Order Pending":
                    driver.find_element_by_xpath('//*[@id="main-container"]/form/div/div/div[1]/table/tbody/tr[{}]/td[10]/input'.format(i)).click()
                    count += 1
        except: pass
        submit = driver.find_element_by_xpath('//*[@id="main-container"]/form/div/div/div[2]/input[1]')
        submit.click()
        print('Successfully clicked {} boxes'.format(count))
        return [0,driver,count]
    except:
        return [1,driver,0]

def capture_payments(driver,orderID,domain):
    try:
        url = 'https://www.'+domain+'/admin/orders/cccharge.cfm?ordernum=' + orderID
        driver.get(url)
        print('    '+orderID + ' - Capturing Payment')
        total = driver.find_element_by_xpath('/html/body/form/div[1]/div[2]/div/table/tbody/tr[4]/td[2]').text
        prev_payments = driver.find_element_by_xpath('/html/body/form/div[1]/div[2]/div/table/tbody/tr[5]/td[2]').text
        due = driver.find_element_by_xpath('/html/body/form/div[1]/div[2]/div/table/tbody/tr[6]/td[2]').text
        print('        Total:    '+str(total))
        print('        Captured: '+str(prev_payments))
        print('        Due:      '+str(due))
        if str(due) == '$0.00':
            print('    No payment captured required')
            return [0,driver,due]
        else: 
            submit = driver.find_element_by_xpath('/html/body/form/div[2]/input[1]')
            submit.click()
            result = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/table/tbody/tr[1]/td').text
            print('Capture Result: '+result)
            if result == 'Congratulations! The transaction was successful.':
                return [0,driver,result]
            else:
                return [2,driver,result]
    except:
        return [1,driver,'0']

def request_fulfillment(driver,orderID,itemIDs,domain,pref_vendor='Trek'):
    try:
        fulfilledIDs = []
        selected_vendor = ''
        url = 'https://www.'+domain+'/admin/orders/requestfulfillment.cfm?orderid='+orderID
        driver.get(url)
        vendor_list = driver.find_elements_by_class_name('seafulfillmenttab')
        for vendor in vendor_list:
            if vendor.get_attribute('class') == "seafulfillmenttab active":
                if vendor.get_attribute("title") == 'Trek - US': selected_vendor = "Trek"
                elif vendor.get_attribute("title") == 'QBP': selected_vendor = "QBP"
                elif vendor.get_attribute("title") == 'HLC - US': selected_vendor = "HLC"
                else:
                    print('    Failed to identify current vendor, aborting')
                    return [2,driver,'']
        if selected_vendor != pref_vendor:
            if pref_vendor == 'Trek':
                try: 
                    driver.find_element_by_id('fulfillment_1').click()
                    time.sleep(0.5)
                    print(driver.find_element_by_id('fulfillment_1').get_attribute("class"))
                    if driver.find_element_by_id('fulfillment_1').get_attribute("class") == 'seafulfillmenttab active':
                        print('    Successfully changed tab to {}'.format(pref_vendor))
                        if driver.find_element_by_id('fulfillment_1').get_attribute("title") == 'Trek - US': selected_vendor = "Trek"
                        elif driver.find_element_by_id('fulfillment_1').get_attribute("title") == 'QBP': selected_vendor = "QBP"
                        elif driver.find_element_by_id('fulfillment_1').get_attribute("title") == 'HLC - US': selected_vendor = "HLC"
                    else:
                        print('    Failed to switch vendor tabs, second tab exists though, aborting fulfillment')
                        return [2,driver,'']
                except:
                    print('    Failed to switch vendor tabs, there may be no second tab, aborting fulfillment')
            elif pref_vendor == 'QBP':
                try: 
                    driver.find_element_by_id('fulfillment_6').click()
                    time.sleep(0.5)
                    print(driver.find_element_by_id('fulfillment_6').get_attribute("class"))
                    if driver.find_element_by_id('fulfillment_6').get_attribute("class") == 'seafulfillmenttab active':
                        print('    Successfully changed tab to {}'.format(pref_vendor))
                        if driver.find_element_by_id('fulfillment_6').get_attribute("title") == 'Trek - US': selected_vendor = "Trek"
                        elif driver.find_element_by_id('fulfillment_6').get_attribute("title") == 'QBP': selected_vendor = "QBP"
                        elif driver.find_element_by_id('fulfillment_6').get_attribute("title") == 'HLC - US': selected_vendor = "HLC"
                    else:
                        print('    Failed to switch vendor tabs, second tab exists though, aborting fulfillment')
                        return [2,driver,'']
                except:
                    print('    Failed to switch vendor tabs, there may be no second tab, aborting fulfillment')
                    return [2,driver,'']
            else:
                print('    Unknown Vendor, aborting')
                return [2,driver,'']
        if selected_vendor == 'Trek': xpath_start = '//*[@id="tab_1"]/form/div[1]/div/table/tbody/tr['
        elif selected_vendor == 'QBP': xpath_start = '//*[@id="tab_6"]/form/div[1]/div/table/tbody/tr['
        checks = 0
        try:
            for i in range(2,15):
                partnoxpath = xpath_start + str(i)+ ']/td[3]'
                partno = driver.find_element_by_xpath(partnoxpath).text
                partno = partno.split(' ')
                partno = partno[len(partno)-2]
                xpath = xpath_start + str(i)+ ']/td[1]/input'
                if str(partno) in itemIDs:
                    driver.find_element_by_xpath(xpath).click()
                    checks =+ 1
                    fulfilledIDs.append(partno)
                    continue
                else: continue
        except:
            print('    End of items, checked {} boxes'.format(checks))
        if checks == 0:
            return [2,driver,'']
        if selected_vendor == 'Trek': xpath = '//*[@id="tab_1"]/form/div[2]/input[1]'
        elif selected_vendor == 'QBP': xpath = '//*[@id="tab_6"]/form/div[2]/input[1]'
        try: 
            driver.find_element_by_xpath(xpath).click()
            driver.find_element_by_xpath('/html/body/form/div[2]/input[1]').click()
            return [0,driver,fulfilledIDs]
        except:
            print('    Failed to Submit Fulfillment')
            return [2,driver,'']
    except:
        print('    Fulfillment page failed')
        return [1,driver,''] 

def change_shipping(driver,orderID,domain):
    try:
        url = 'https://www.'+domain+'/admin/orders/editorder.cfm?orderid='+orderID
        driver.get(url)
        driver.find_element_by_xpath('//*[@id="shippingMethodSelect"]/option[text()="Ship - Ground"]').click()
        driver.find_element_by_xpath('//*[@id="updateallForm"]/div[2]/input[2]').click()
        print('    Shipping Method for order ID '+orderID+' updated to "Ship - Ground"')
        return [0,driver]
    except:
        print('    Update Shipping method faild for orderID '+orderID)
        return [1,driver]

def get_discount_ID(driver,orderID,domain):
    try:
        url = 'https://www.'+domain+'/admin/index.cfm?OrderID='+orderID
        driver.get(url)
        discount = driver.find_element_by_xpath('//*[@id="main-container"]/div[1]/div/div[2]/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr[2]/td[2]').text
        return [0,driver,discount]
    except:
        print('    FAILED to get discount ID')
        return [1,driver,'Unknown']

def get_payment_ID(driver,orderID,domain):
    try:
        url = 'https://www.'+domain+'/admin/index.cfm?OrderID='+orderID
        driver.get(url)
        paymentID = driver.find_element_by_xpath('//*[@id="main-container"]/div[1]/div/div[2]/table/tbody/tr[6]/td/table/tbody/tr[2]/td[4]').text
        status = driver.find_element_by_xpath('//*[@id="main-container"]/div[1]/div/div[2]/table/tbody/tr[6]/td/table/tbody/tr[2]/td[5]').text
        return [0,driver,paymentID,status]
    except:
        print('    FAILED to get payment ID')
        return [1,driver,'Unknown','Unknown']

def update_shipping_cost(driver,domain):
    errors = 0
    try:
        pages = []
        url = 'https://www.'+domain+'/admin/items/listitem.cfm'
        driver.get(url)
        wait = WebDriverWait(driver, 30)
        wait.until(ec.visibility_of_element_located((By.XPATH, ('//*[@id="catalogspopuplink"]'))))
        driver.find_element_by_xpath('//*[@id="catalogspopuplink"]').click()
        wait.until(ec.visibility_of_element_located((By.XPATH, ('//*[@id="CategoriesSelect"]'))))
        iframe = driver.find_element_by_xpath('//*[@id="CategoriesSelect"]')
        driver.switch_to.frame(iframe)
        wait.until(ec.visibility_of_element_located((By.XPATH, ('//*[@id="catalogtree"]/ul/li[1]/div[1]'))))
        driver.find_element_by_xpath('//*[@id="catalogtree"]/ul/li[1]/div[1]').click()
        wait.until(ec.visibility_of_element_located((By.XPATH, ('//*[@id="c1000"]'))))
        driver.find_element_by_xpath('//*[@id="c1000"]').click()
        driver.find_element_by_xpath('//*[@id="c1289"]').click()
        driver.find_element_by_xpath('//*[@id="catalogtree"]/ul/li[1]/ul/li[7]/div[1]').click()
        wait.until(ec.visibility_of_element_located((By.XPATH, ('//*[@id="c1264"]'))))
        driver.find_element_by_xpath('//*[@id="c1264"]').click()
        driver.find_element_by_xpath('//*[@id="c1217"]').click()
        driver.find_element_by_xpath('//*[@id="c1218"]').click()
        driver.find_element_by_xpath('//*[@id="c1219"]').click()
        driver.find_element_by_xpath('//*[@id="c1220"]').click()
        wait.until(ec.visibility_of_element_located((By.XPATH, ('//*[@id="secontrols"]/input[1]'))))
        driver.find_element_by_xpath('//*[@id="secontrols"]/input[1]').click()
        driver.switch_to.default_content()
        wait.until(ec.element_to_be_clickable((By.XPATH, ('//*[@id="main-container"]/form/div[1]/div[2]/div/table/tbody/tr[1]/td[2]/table/tbody/tr[2]/td[2]/input'))))
        time.sleep(2)
        driver.find_element_by_xpath('//*[@id="main-container"]/form/div[1]/div[2]/div/table/tbody/tr[1]/td[2]/table/tbody/tr[2]/td[2]/input').click()
        driver.find_element_by_xpath('//*[@id="main-container"]/form/div[2]/input[1]').click()
        count = 1
        wait.until(ec.visibility_of_element_located((By.XPATH, ('//*[@id="main-container"]/div[1]/form[2]/div[1]/div[2]/div[1]/div/ul/li[2]/a'))))
        while True:
            page_item_list=[]
            page_item_list=driver.find_elements_by_class_name('seaicon_edit')
            if len(page_item_list) > 0:
                for item in page_item_list:
                    pages.append(item.get_attribute('href'))
                try:
                    if count == 1:
                        count = 2
                        driver.find_element_by_xpath('//*[@id="main-container"]/div[1]/form[2]/div[1]/div[2]/div[4]/ul/li[3]/a').click()
                    elif count == 2:
                        count = 3
                        driver.find_element_by_xpath('//*[@id="main-container"]/div[1]/form[2]/div[1]/div[2]/div[4]/ul/li[4]/a').click()
                    elif count == 3:
                        count = 4
                        driver.find_element_by_xpath('//*[@id="main-container"]/div[1]/form[2]/div[1]/div[2]/div[4]/ul/li[5]/a').click()
                    else:
                        count = 5
                        driver.find_element_by_xpath('//*[@id="main-container"]/div[1]/form[2]/div[1]/div[2]/div[4]/ul/li[6]/a').click()
                except: break
            else:break
        
        for page in pages:
            try:
                driver.get(page)
                wait.until(ec.visibility_of_element_located((By.XPATH, ('//*[@id="3"]/a'))))
                driver.find_element_by_xpath('//*[@id="3"]/a').click()
                wait.until(ec.visibility_of_element_located((By.XPATH, ('//*[@id="tab1-3"]/div[2]/div/div[2]/div[2]/div/table/tbody/tr[3]/td[10]/input'))))
                price = float(driver.find_elements_by_class_name('validatePrice')[0].get_attribute('Value'))
                print(price)
                category = ''
                isbike = False
                z = re.search('Bikes',driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/form/div[2]/div/table/tbody/tr[3]/td[2]/div[1]/div/span').text)
                if z is None: 
                    description = re.search('Car Racks',driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/form/div[2]/div/table/tbody/tr[3]/td[2]/div[1]/div/span').text)
                    if description is None: continue
                    else: category = 'Auto'
                else:
                    description = re.search('Frame',driver.find_element_by_xpath('//*[@id="itemForm"]/div[2]/div/table/tbody/tr[2]/td[2]/div[1]/input').get_attribute('value'))
                    if description is None:
                        categories = driver.find_elements_by_xpath('//*[@id="itemcategories"]/div')
                        for catitem in categories:
                            ebike = re.search('Electric',catitem.text)
                            if ebike is None: continue
                            category = 'Ebike'
                        if category != 'Ebike':
                            category = 'Bike'
                            isbike = True
                    else:
                        category = 'Frame'
                if category == 'Ebike': shipping = 189
                elif category == 'Auto': shipping = 49
                elif category == 'Frame': shipping = 89
                elif isbike == True and price < 800: shipping = 89
                elif isbike == True and price < 3000: shipping = 129
                elif isbike == True: shipping = 159
                print(category)
                print(shipping)
                
                wait.until(ec.visibility_of_element_located((By.XPATH,('//*[@id="4"]/a'))))
                driver.find_element_by_xpath('//*[@id="4"]/a').click()
                wait.until(ec.visibility_of_element_located((By.NAME, ('CHARGESHIPPING'))))
                time.sleep(1)
                if driver.find_element_by_name('CHARGESHIPPING').is_selected() is False:
                    driver.find_element_by_name('CHARGESHIPPING').click()
                wait.until(ec.visibility_of_element_located((By.NAME, ('SHIPCHARGE'))))
                driver.find_element_by_name('SHIPCHARGE').click()
                driver.find_element_by_name('SHIPCHARGE').clear()
                time.sleep(0.5)
                driver.find_element_by_name('SHIPCHARGE').send_keys(shipping)
                time.sleep(0.5)
                driver.find_element_by_name('SHIPCHARGE').send_keys(Keys.TAB)
                time.sleep(0.5)
                driver.find_element_by_name('SHIPCHARGE').send_keys(Keys.ENTER)
                time.sleep(1)
                wait.until(ec.element_to_be_clickable((By.NAME,('Submit'))))
                # driver.find_element_by_name('Submit').click()
                time.sleep(5)
            except:
                errors += 1
                print('ERROR # {}'.format(errors))
        
            
    except:
        print(traceback.format_exc())
        return [1,driver,'Unknown','Unknown']

def create_unique_discount(driver,DiscountName,DiscountCode,DiscountMethod,DiscountAmount,ExpireDate,domain):
    try:
        url = 'https://www.'+domain+'/admin/discounts/editmethod.cfm?add=true'
        driver.get(url)
        driver.find_element_by_xpath('//*[@id="newDiscountMethod"]/div[1]/div/table/tbody/tr[1]/td[2]/div/input').send_keys(DiscountName)
        driver.find_element_by_xpath('//*[@id="newDiscountMethod"]/div[1]/div/table/tbody/tr[2]/td[2]/div/input').send_keys(DiscountCode)
        driver.find_element_by_xpath('//*[@id="newDiscountMethod"]/div[1]/div/table/tbody/tr[3]/td[2]/label/input').click()
        driver.find_element_by_xpath('//*[@id="newDiscountMethod"]/div[2]/input[2]').click()
        driver.find_element_by_xpath('//*[@id="main-container"]/div[1]/input').click()
        driver.find_element_by_xpath('//*[@id="addDiscountRule"]/div[1]/div/table/tbody/tr[1]/td[2]/div/input[1]').send_keys('1')
        driver.find_element_by_xpath('//*[@id="addDiscountRule"]/div[1]/div/table/tbody/tr[1]/td[2]/div/input[2]').send_keys('1500')
        driver.find_element_by_xpath('//*[@id="addDiscountRule"]/div[1]/div/table/tbody/tr[2]/td[2]/div/input').send_keys(DiscountAmount)
        driver.find_element_by_xpath('//*[@id="addDiscountRule"]/div[1]/div/table/tbody/tr[2]/td[2]/div/select/option[text()="{}"]'.format(DiscountMethod)).click()
        driver.find_element_by_xpath('//*[@id="addDiscountRule"]/div[1]/div/table/tbody/tr[3]/td[2]/div/input').send_keys('1')
        driver.find_element_by_xpath('//*[@id="addDiscountRule"]/div[1]/div/table/tbody/tr[3]/td[2]/div/select/option[2]').click()
        driver.find_element_by_xpath('//*[@id="addDiscountRule"]/div[2]/div/div/input[1]').click()
        driver.find_element_by_xpath('//*[@id="updateDiscountMethod"]/div[1]/div/table/tbody/tr[2]/td[2]/label/input').click()
        driver.find_element_by_xpath('//*[@id="enddate"]').send_keys(ExpireDate)
        driver.find_element_by_xpath('//*[@id="updateDiscountMethod"]/div[2]/input[3]').click()
        return [0, driver]
    except:
        print('    Failed to create Unique code!')
        print(traceback.format_exc())
        return [1, driver]

def delete_old_discounts(driver,domain):
    wait = WebDriverWait(driver, 30)
    count = 0
    try:
        count = 0
        row = 3
        while True:
            url = 'https://www.'+domain+'/admin/discounts/index.cfm?startpage=1&currpage=11&startrow=501'
            driver.get(url)
            try:
                status = driver.find_element_by_xpath('//*[@id="main-container"]/form/div[2]/div/table/tbody/tr[{}]/td[3]/span'.format(row)).text
            except:
                return [0,driver,count]
            if status == 'Expired':
                driver.find_element_by_xpath('//*[@id="main-container"]/form/div[2]/div/table/tbody/tr[{}]/td[1]/a[2]/i'.format(row)).click()
                time.sleep(1)
                driver.switch_to.active_element
                wait.until(ec.element_to_be_clickable((By.XPATH,('/html/body/div[6]/div/div/div[2]/button[2]'))))
                driver.find_element_by_xpath('/html/body/div[6]/div/div/div[2]/button[2]').click()
            else: row += 1
        return [0,driver,count]
    except: 
        print(traceback.format_exc())
        return [1,driver,count]

def get_product_map(driver,domain):
    try:
        url = 'https://www.'+domain+'/admin/items/listitem.cfm'
        driver.get(url)
        driver.find_element_by_xpath('//*[@id="main-container"]/form/div[2]/input[1]').click()
        driver.find_element_by_xpath('//*[@id="main-container"]/div[1]/form[1]/div[2]/input[2]').click()
        # Check all boxes of options
        boxnumbers = [1,7,8,9,11,12,13,14,15,16,17,18,19,20,21,22]
        for box in boxnumbers:
            driver.find_element_by_xpath('//*[@id="formItemDownload"]/div[1]/div[2]/table/tbody/tr[2]/td[2]/table/tbody/tr[{}]/td[1]/input'.format(box)).click()
        driver.find_element_by_xpath('//*[@id="formItemDownload"]/div[2]/input[31]').click()
        time.sleep(60)
        return [0,driver]
    except:
        print('Failed to get SE Product Map File')
        print(traceback.format_exc())
        return [1,driver]

def update_item_notes(driver,itemID,note,domain):
    url = 'https://www.'+domain+'/admin/orders/orderdetails_note.cfm?id={}'.format(itemID)
    driver.get(url)
    t = driver.find_element_by_xpath('/html/body/div[2]/form/div/div/table/tbody/tr[1]/td/div/textarea')
    current = t.text
    t.click()
    t.clear()
    t.send_keys(current)
    t.send_keys(Keys.ENTER)
    t.send_keys(note)
    driver.find_element_by_xpath('/html/body/div[2]/form/div/div/div/input[1]').click()
    return [0,driver]
    
def update_ordernotes(driver,OrderID,note,domain):
    url = 'https://www.'+domain+'/admin/orders/vieworder.cfm?OrderID={}'.format(OrderID)
    driver.get(url)
    t = driver.find_element_by_name('Comment')
    current = t.text
    t.click()
    t.clear()
    t.send_keys(current)
    t.send_keys(Keys.ENTER)
    t.send_keys(note)
    driver.find_element_by_name('submit').click()
    return [0,driver]


def get_config():
    with open('config.yaml') as file:
        config = yaml.load(file, Loader=yaml.FullLoader)
    return config
