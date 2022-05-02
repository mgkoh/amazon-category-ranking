import os
import pathlib
import re
from time import sleep
import time
from datetime import datetime
import Screenshot.Screenshot_Clipping       #pip install Selenium-Screenshot
import selenium
import xlsxwriter
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By


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
            ASIN_xpath = browser.find_element(By.XPATH, '//*[@id="productDetails_detailBullets_sections1"]/tbody/tr[1]/td').get_attribute("innerHTML").strip(" ")
            price_ID = browser.find_element(By.CSS_SELECTOR, '.a-price .a-offscreen').get_attribute("innerHTML").strip('£')

        except NoSuchElementException:
            price_ID = "Not Available"

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
                if (quantity.find("count")!= -1):
                    quantity= quantity.strip("count")
                elif(quantity.find("gram")!= -1 or quantity.find("grams")!= -1):
                    quantity = quantity.replace("gram", "g")
                elif(quantity.find("millilitre")!= -1 or quantity.find("milliliter")!= -1):
                    quantity = quantity.replace("millilitre" or "milliliter", "ml")
            else:
                quantity = ""


            if(".00" in quantity):
                quantity = quantity.replace(".00","")
            elif(".0" in quantity):
                quantity = quantity.replace(".0", "")


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
        for item in product:
            if (item.find("amp;") !=-1):
                item = item.strip("amp;")


        details.append(product)

        browser.close()
        browser.switch_to.window(window_list[0])
        print(len(details))

def writing_excel(section_name, product_details, img_names):
    ###writing excel file
    row = 0
    col = 0

    excel_format = ["Type","Ranking", "ASIN", "Product Name", "Qty", "Capsules/Tablets", "Brand", "Spec", "RRP Total",  "Main Category Rank", "URL", "Date", "Sales Qty"]   ###Excel format title
    # Create a workbook and add a worksheet.
    workbook_ranking = xlsxwriter.Workbook(datetime.today().strftime('%Y-%m-%d') + " " + section_name + " Ranking"+ '.xlsx')
    worksheet_ranking = workbook_ranking.add_worksheet(section_name[:30])
    worksheet_screenshot = workbook_ranking.add_worksheet("Screenshots")

    ###writing format for excel format
    for format in excel_format:
        worksheet_ranking.write(row, col, format)
        col += 1

    print("done writing excel format!\n")
    row = 1
    col = 0

    ####
    ####Excel writing contents
    ####
    #    product=[title_xpath, ASIN_xpath, price_ID, brand_name, url, dosage_form, qty, main_category_ranking]  ###proudct list items
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

        worksheet_ranking.write(row, col, datetime.today().strftime('%Y-%m-%d')) #date
        col = col +1

        worksheet_ranking.write(row, col, "") #Sales Qty



        row+=1


    ###
    ###add new worksheet and screenshots
    ###
    worksheet_screenshot.write(0,0, datetime.today().strftime('%Y-%m-%d')+'T'+datetime.now().strftime('%H:%M'))
    count = 0
    for path in img_names:
        if(count==0):
            cell = 'B2'
            worksheet_screenshot.insert_image(cell, path)
            count = count+1

        else:
            cell = 'AI2'
            worksheet_screenshot.insert_image(cell, path)

        # os.remove(str(pathlib.Path(__file__).parent.resolve())+'/'+path)

    print("Done : "+section_name)

    workbook_ranking.close()


    #deletes screenshots
    for path in img_names:
        os.remove(str(pathlib.Path(__file__).parent.resolve()) + '/'+ path)

def ranking_page_scroll():
    browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    sleep(1)


    try:
        #capture and scroll to item 30
        element_30 = browser.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[1]/div[1]/div[31]/div/div[2]/div/a[1]')
        browser.execute_script("arguments[0].scrollIntoView();", element_30)
        sleep(1.5)

        #capture and scroll to item 38
        element_38 = browser.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[1]/div[1]/div[39]/div/div[2]/div/a[1]')
        browser.execute_script("arguments[0].scrollIntoView();", element_38)
        sleep(1.5)

        #capture and scroll to item 46
        element_46 = browser.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[1]/div[1]/div[47]/div/div[2]/div/a[1]')
        browser.execute_script("arguments[0].scrollIntoView();", element_46)
        sleep(1.5)

        #removes cookie permission element
        cookie_element = browser.find_element(By.XPATH, '//*[@id="sp-cc"]')
        browser.execute_script("""var element = arguments[0]; element.parentNode.removeChild(element);""", cookie_element)

    except (selenium.common.exceptions.NoSuchElementException):
        print("might not have 46th or 96th product in the page.")
        pass

    finally:

        ###
        ###retrieves ranking page title and screenshots the page
        ###
        section_name = (browser.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[1]/h1').get_attribute("innerHTML")).replace("Best Sellers in", "").strip(" ").replace("&amp;", "&")


        img_name = datetime.today().strftime('%Y-%m-%d')+'T'+datetime.now().strftime('%H-%M')+'-'+section_name+'.png'

        ob = Screenshot.Screenshot_Clipping.Screenshot()

        ###image is saved in the same directory with category name, will be removed when operation is completed
        img = ob.full_Screenshot(browser, save_path=str(pathlib.Path(__file__).parent.resolve()), image_name=img_name)
        print(img_name+" has been saved.")
        return img_name, section_name

def product_url_list():
    # ranking page collecting url using href
    products_grid_xpath = '//*[@id="gridItemRoot"]/div/div/div/a[1]'
    products = browser.find_elements(By.XPATH, products_grid_xpath)
    lists = [element.get_attribute('href') for element in products]
    return lists

start = time.time()
try:
    with open("links" + '.txt') as f:
        links = f.read().splitlines()
    for link in links:
        browser = webdriver.Chrome()
        browser.get(link)
        browser.maximize_window()
        img_names =[]
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


        img_name, section_name = ranking_page_scroll()
        img_names.append(img_name)


        details = []

        get_product_info(details, product_url_list())

        ###
        ###acessing to second page for item #51 to #100
        ###

        try:
            next_page_button = browser.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[2]/ul/li[4]')
            browser.execute_script("arguments[0].scrollIntoView();", next_page_button)
            sleep(1.5)
            next_page_button.click()
            img_name, section_name = ranking_page_scroll()
            img_names.append(img_name)
            get_product_info(details,product_url_list())

        except(NoSuchElementException):
            print("Next page button not found on ranking page.")

        finally:

            browser.quit()
            writing_excel(section_name, details, img_names)


except(FileNotFoundError):
    print("links.txt is not found.")

except(selenium.common.exceptions.InvalidArgumentException):
    print("Error: link is null or can't be reached.")

finally:
    ###
    ###prints elapsed time
    ###
    end = time.time()
    print("Total Elapsed Time :"+str(end - start) +"s")