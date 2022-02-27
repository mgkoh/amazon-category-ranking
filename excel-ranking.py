import re
from time import sleep
from datetime import datetime
import xlsxwriter
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

def get_product_info(details, lists):
    for url in lists:
        browser.execute_script("window.open('');")
        window_list = browser.window_handles
        sleep(0.5)
        browser.switch_to.window(window_list[-1])
        browser.get(url)

    #   #details in product page
        sleep(0.5)
        try:
            title_xpath = browser.find_element(By.XPATH, '//*[@id="productTitle"]').get_attribute("innerHTML").strip("        ")
            table = browser.find_elements(By.CSS_SELECTOR, 'table.a-keyvalue td, table.a-keyvalue th')

            # dosage_form_xpath = browser.find_element(By.XPATH, '//*[@id="productOverview_feature_div"]/div/table/tbody/tr[6]/td[2]/span').get_attribute("innerHTML")
            ASIN_xpath = browser.find_element(By.XPATH, '//*[@id="productDetails_detailBullets_sections1"]/tbody/tr[1]/td').get_attribute("innerHTML").strip(" ")

            price_ID = browser.find_element(By.CSS_SELECTOR, '.a-price .a-offscreen').get_attribute("innerHTML")
        except NoSuchElementException:
            price_ID = "Not Available"

        # print(title_xpath)
        # print(dosage_form_xpath)
        # print(ASIN_xpath)
        # print(price_ID)
        table_str = []



        for item in table:
            table_str.append(item.get_attribute("innerHTML"))

        try:

            if((" Brand " or "Brand" in item for item in table_str)):
                brand_name = table_str[table_str.index(" Brand ") + 1].strip("\n                \u200e")
            elif((" Manufacturer " or "Manufacturer" in item for item in table_str)):
                brand_name = table_str[table_str.index(" Manufacturer ") + 1].strip("\n                \u200e")
            else:
                brand_name = ""

            if('&' in brand_name):
                brand_name = brand_name.replace("&amp;", "&")
                title_xpath = title_xpath.replace("&amp;", "&")

            if (any(" Units " in item for item in table_str)):
                quantity = table_str[table_str.index(" Units ") + 1].strip("\n                \u200e")
                if ("count" in quantity):
                    quantity.strip(".00 count")
            else:
                quantity = ""

        except (ValueError, UnboundLocalError) as e:
            brand_name = ""
            quantity = ""
            pass

        if(any(" Format " in item for item in table_str)):
            dosage_form = table_str[table_str.index(" Format ") + 1].strip("\n                \u200e")
        elif("Tablet" or "Tablets" in title_xpath):
            dosage_form = "Tablet"
        elif("Capsule" or "Capsules" in title_xpath):
            dosage_form = "Capsule"
        else:
            dosage_form = ""



        if(any(" Best Sellers Rank " in item for item in table_str)):
            category_rank = table_str[table_str.index(" Best Sellers Rank ") + 1].strip(" <span>  <span>")
            category_rank= category_rank[:8]
            category_rank = re.sub("[^0-9]","",category_rank)

        # brand_name =""
        link = "https://amazon.co.uk/dp/"+ASIN_xpath
        product=[title_xpath, ASIN_xpath, price_ID, brand_name, link, dosage_form, quantity, category_rank]
        details.append(product)

        browser.close()
        browser.switch_to.window(window_list[0])
        print(len(details))

def writing_excel(section_name, product_details):
    ###writing excel file
    row = 0
    col = 0

    excel_format = ["Type","Ranking", "ASIN", "Product Name", "Qty", "Capsules/Tablets", "Brand", "Spec", "RRP Total",  "Main Category Rank", "URL", "Top Rank (sub)", "Sales Qty"]
    # Create a workbook and add a worksheet.
    workbook_ranking = xlsxwriter.Workbook(datetime.today().strftime('%Y-%m-%d') + " " + section_name + " Ranking" + '.xlsx')
    worksheet_ranking = workbook_ranking.add_worksheet()

    ###writing format for excel format
    for format in excel_format:
        worksheet_ranking.write(row, col, format)
        col += 1

    print("done writing excel format!\n")
    row = 1
    col = 0
    #    product=[title_xpath, ASIN_xpath, price_ID, brand_name, url, dosage_form, qty, main_category_ranking]
    for items in product_details:
        col = 0
        worksheet_ranking.write(row, col, section_name) #type
        col = col +1

        worksheet_ranking.write(row, col, product_details.index(items)+1) #ranking
        col = col +1

        worksheet_ranking.write(row, col, items[1]) #ASIN
        col = col +1

        worksheet_ranking.write(row, col, items[0]) #Product Name
        col = col +1

        worksheet_ranking.write(row, col, items[6]) #Qty
        col = col +1

        worksheet_ranking.write(row, col, items[5]) #capsules/tablets
        col = col +1

        worksheet_ranking.write(row, col, items[3]) #Brand
        col = col +1

        worksheet_ranking.write(row, col, "") #Spec
        col = col +1

        worksheet_ranking.write(row, col, items[2]) #Price
        col = col +1

        worksheet_ranking.write(row, col, items[7]) #main category ranking
        col = col +1

        worksheet_ranking.write(row, col, items[4]) #URL
        col = col +1

        worksheet_ranking.write(row, col, product_details.index(items)+1) #ranking
        col = col +1

        worksheet_ranking.write(row, col, "") #Sales Qty



        row+=1
    print("Done : "+section_name)

    workbook_ranking.close()

def ranking_page_scroll():
    browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    sleep(1)
    element_30 = browser.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[1]/div[1]/div[31]/div/div[2]/div/a[1]')
    browser.execute_script("arguments[0].scrollIntoView();", element_30)
    sleep(1.5)
    element_38 = browser.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[1]/div[1]/div[39]/div/div[2]/div/a[1]')
    browser.execute_script("arguments[0].scrollIntoView();", element_38)
    sleep(1.5)
    element_46 = browser.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[1]/div[1]/div[47]/div/div[2]/div/a[1]')
    browser.execute_script("arguments[0].scrollIntoView();", element_46)
    sleep(1.5)

def product_url_list():
    # ranking page collecting url
    products_grid_xpath = '//*[@id="gridItemRoot"]/div/div/div/a[1]'
    products = browser.find_elements(By.XPATH, products_grid_xpath)
    lists = [element.get_attribute('href') for element in products]
    return lists


with open("links2" + '.txt') as f:
    links = f.read().splitlines()
for link in links:
    browser = webdriver.Chrome()
    browser.get(link)
    browser.maximize_window()


    #change address
    deliver_button_xpath = '//*[@id="nav-global-location-popover-link"]'
    postcode_box_xpath = '//*[@id="GLUXZipUpdateInput"]'
    apply_button_xpath = '//*[@id="GLUXZipUpdate"]/span/input'


    browser.find_element(By.XPATH, deliver_button_xpath).click()
    sleep(1)
    browser.find_element(By.XPATH, postcode_box_xpath).send_keys("NE1 1EE")
    browser.find_element(By.XPATH, apply_button_xpath).click()
    sleep(1.5)

    browser.refresh()
    sleep(3)

    ranking_page_scroll()

    section_name = (browser.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[1]/h1').get_attribute("innerHTML")).replace("Best Sellers in", "").strip(" ")

    details = []

    get_product_info(details, product_url_list())

    next_page_button = browser.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[2]/ul/li[4]')
    browser.execute_script("arguments[0].scrollIntoView();", next_page_button)
    sleep(1.5)
    next_page_button.click()
    ranking_page_scroll()
    get_product_info(details,product_url_list())

    browser.quit()
    writing_excel(section_name, details)