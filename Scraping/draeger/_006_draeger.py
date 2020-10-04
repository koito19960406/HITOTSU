#!/usr/bin/env python
# coding: utf-8

# # Import Libraries

# In[5]:


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import datetime
import csv
import openpyxl
from openpyxl import Workbook
import pandas as pd


# # Create Functions

# In[53]:


class Draeger:
    # prepare the input date
    def __init__(self, year, month, day):
        self.output_date=int(str(year)+str(month)+str(day))

    # scrape info about: name of the product and url
    def scrape(self):
        #open chrome in incognito mode
        options = webdriver.ChromeOptions()
        options.add_argument(' -- incognito')
        browser = webdriver.Chrome(chrome_options=options)

        # deal with the first "medical staff?" question
        browser.get('https://www.draeger.com/ja_jp/Hospital/Portfolio')

        # wait for browser to open for 10 sec
        timeout = 10
        try:
            WebDriverWait(browser, timeout).until(
            EC.visibility_of_element_located(
            (By.XPATH, '//*[@id="acceptdisclaimer_ja_jp_hospital"]')
            )
            )
        except TimeoutException:
            print('Timed Out Waiting for page to load')
            browser.quit()

        # Click the yes button
        login_btn=browser.find_element_by_xpath('//*[@id="acceptdisclaimer_ja_jp_hospital"]')
        login_btn.click()
        browser.implicitly_wait(3)

        # Get info
        # Get the URL list except for accessories and consumables
        page_list=browser.find_elements_by_css_selector("div[class='carditem ng-scope']")[:8]
        page_url_list= [page.find_element_by_css_selector('figure>a').get_attribute('href') for page in page_list]

        # create a list
        result=[]
        # Go through the url list
        for page_url in page_url_list:
            browser.get(page_url)

            # multiple lists page
            try: 
                '''outer-most variables'''
                product_category_url_raw_list=browser.find_element_by_id('tabsection_0').find_elements_by_css_selector(
                    "li[class='navitem ng-scope']>a")
                product_category_url_list=[url_raw.get_attribute('href') for url_raw in product_category_url_raw_list]

                ''''product category layer'''
                for i, url in enumerate(product_category_url_list):
                    browser.get(url)
                    try:
                        WebDriverWait(browser, timeout).until(
                        EC.visibility_of_element_located(
                        (By.XPATH, '//*[@id="tabsection_0"]')
                        )
                        )
                    except TimeoutException:
                        print('Timed Out Waiting for page to load')
                        browser.quit()
                    browser.implicitly_wait(3)  
                    product_category_list=browser.find_element_by_id('tabsection_0').find_elements_by_class_name('tabsectionpanel')
                    html_content=product_category_list[i].get_attribute('innerHTML')

                    # for listed category
                    try:
                        swiper_wrapper=product_category_list[i].find_element_by_css_selector("div[class='swiper-wrapper']")
                        product_name_list=[]
                        product_url_list=[]
                        product_url_temp_list=product_category_list[i].find_elements_by_css_selector(
                            "a[class='primarybtn small ng-binding ng-scope ng-isolate-scope']")
                        for product_url_raw in product_url_temp_list:
                            product_url=product_url_raw.get_attribute('href')
                            product_url_list.append(product_url)
                        
                        '''InnerHTML'''
                        browser.get("data:text/html;charset=utf-8,{html_content}".format(html_content=html_content))
                        product_list=browser.find_elements_by_css_selector('div.swiper-wrapper>div')
                        for product in product_list:
                            product_name=product.find_element_by_css_selector('header>h3').text.strip()
                            product_name_list.append(product_name)

                        for product_name, product_url in zip(product_name_list,product_url_list):
                            result.append([product_name, product_url])

                    # for non-listed category
                    except NoSuchElementException:
                        product_name=product_category_list[i].find_element_by_css_selector(
                            "h3[class='ng-binding ng-scope']").text.strip()
                        product_url=product_category_list[i].find_element_by_css_selector(
                            "a[class='primarybtn small ng-binding ng-isolate-scope']").get_attribute('href')
                        result.append([product_name,product_url])

            # single item page  
            except NoSuchElementException:
                try:
                    product=browser.find_element_by_css_selector(
                        "div[class='contentpart-content twocolumnpart twocolumnformcontentpart']"
                        ).find_element_by_css_selector("div[class='contentbox text']")
                    product_name=product.find_element_by_css_selector("h3[class='ng-binding ng-scope']").text.strip()
                    product_url=product.find_element_by_css_selector(
                        "a[class='primarybtn small ng-binding ng-isolate-scope']").get_attribute('href')
                    result.append([product_name,product_url])

                # single bullet point list
                except NoSuchElementException:
                    product_category=browser.find_element_by_css_selector(
                        "div[class='swiper-container swiper-portfolio swiper-container-horizontal']")
                    html_content=product_category.get_attribute('innerHTML')
                    product_name_list=[]
                    product_url_list=[]
                    product_url_temp_list=product_category.find_elements_by_css_selector(
                        "a[class='primarybtn small ng-binding ng-scope ng-isolate-scope']")
                    for product_url_raw in product_url_temp_list:
                        product_url=product_url_raw.get_attribute('href')
                    product_url_list.append(product_url)

                    '''InnerHTML'''
                    browser.get("data:text/html;charset=utf-8,{html_content}".format(html_content=html_content))
                    product_list=browser.find_elements_by_css_selector('div.swiper-wrapper>div')
                    for product in product_list:
                        product_name=product.find_element_by_css_selector('header>h3').text.strip()
                        product_name_list.append(product_name)

                    for product_name, product_url in zip(product_name_list,product_url_list):
                        result.append([product_name, product_url])



        # close the browser        
        browser.quit()
        return result

    # store the list into CSV if there is no csv. If not, load the csv and check if there's any product thats not in the older csv.
    def get_new_product(self):
        try:
        # load an old list of product and convert it to a dataframe
            column_names=['product_name','product_url']
            df_product_list=pd.read_csv("Draeger_product_list.csv",names=column_names)

            # get result
            result = self.scrape()
            

            # Find if there's any new products
            new_product_list=[]
            for product in result:
                if not(df_product_list['product_name'].isin([product[0]]).any()):
                    new_product_list.append(product)
            
            # check if there's any new product. If so, append them to the existing product list
            if len(new_product_list)>0:
                with open('Draeger_product_list.csv', 'a') as csvfile:
                    writer = csv.writer(csvfile)
                    for i in range(len(new_product_list)):
                        writer.writerow([new_product_list[i][0],  
                                        new_product_list[i][1]])

            # return a list of new products
            return new_product_list
        
        # if there's no such file, create a new file 
        except FileNotFoundError:
            with open('Draeger_product_list.csv','w') as csvfile:
                result = self.scrape()
                result_len=len(result)
                writer = csv.writer(csvfile)
                for i in range(result_len):
                    writer.writerow([result[i][0],  
                                    result[i][1]])
            # return an empty list
            return []

   
   
    # store new products' info into csv
    def to_csv(self):
        # get new products
        new_product_list=self.get_new_product()
        new_product_list_len = len(new_product_list)

        # check if the result is empty
        if new_product_list_len == 0:
            return

        # get row number
        # try to open the csv file
        try:
            with open('product_info.csv') as csvfile:
                reader = csv.reader(csvfile)
                for row in reader:
                    # the date is already there, dont add anything
                    if (row[0]==self.output_date) and (row[4]==6):
                        return print('Already added to csv')

        # if there's no such file, create a new file 
        except FileNotFoundError:
            with open('product_info.csv','w') as csvfile:
                pass
        
        # add new data
        with open('product_info.csv', 'a') as csvfile:
            writer = csv.writer(csvfile)
            for i in range(new_product_list_len):
                writer.writerow([self.output_date, 
                                '人工呼吸器', 
                                '手術室', 
                                '',
                                6,
                                'ドレーゲル・メディカルジャパン', 
                                '新製品',
                                new_product_list[i][0], 
                                new_product_list[i][1], 
                                1])


# # Run the Function

# In[56]:


if __name__=='__main__':
    year=2020
    month=10
    day=4
    draeger=Draeger(year,month,day)
    draeger.to_csv()


# In[ ]:




