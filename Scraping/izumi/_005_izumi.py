#!/usr/bin/env python
# coding: utf-8

# # Import Libraries

# In[82]:


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

# In[85]:


class Izumi:
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
        browser.get('https://www.mera.co.jp/medical/?redirect_url=https://www.mera.co.jp/medical/product-info/')

        # wait for browser to open for 10 sec
        timeout = 10
        try:
            WebDriverWait(browser, timeout).until(
            EC.visibility_of_element_located(
            (By.XPATH, '/html/body/section/section/div[2]/div/div/div/div[2]/a[2]')
            )
            )
        except TimeoutException:
            print('Timed Out Waiting for page to load')
            browser.quit()

        # Click the yes button
        login_btn=browser.find_element_by_xpath('/html/body/section/section/div[2]/div/div/div/div[2]/a[2]')
        login_btn.click()
        browser.implicitly_wait(3)

        # Get info
        # Get the URL list
        page_list=browser.find_elements_by_css_selector('div.c-product-info__wrap>a')
        page_url_list= [page.get_attribute('href') for page in page_list]

        # create a list
        result=[]
        # Go through the url list
        for page_url in page_url_list:
            # go to the URL
            browser.get(page_url)

            # get the product list
            product_list=browser.find_elements_by_css_selector('div.c-box__block')

            # go through the list
            for product in product_list:
                # Get link and title
                try:
                    product_url=product.find_element_by_css_selector("a[class='c-button is-sm']").get_attribute('href')
                except NoSuchElementException:
                    product_url=''
                product_name=product.find_element_by_css_selector("div[class='c-heading is-sm']").text.strip()
                # Append the info to the list
                result.append([product_name, product_url])

        # close the browser        
        browser.quit()
        return result

    # store the list into CSV if there is no csv. If not, load the csv and check if there's any product thats not in the older csv.
    def get_new_product(self):
        try:
        # load an old list of product and convert it to a dataframe
            column_names=['product_name','product_url']
            df_product_list=pd.read_csv("Izumi_product_list.csv",names=column_names)

            # get result
            result = self.scrape()
            

            # Find if there's any new products
            new_product_list=[]
            for product in result:
                if not(df_product_list['product_name'].isin([product[0]]).any()):
                    new_product_list.append(product)
            
            # check if there's any new product. If so, append them to the existing product list
            if len(new_product_list)>0:
                with open('Izumi_product_list.csv', 'a') as csvfile:
                    writer = csv.writer(csvfile)
                    for i in range(len(new_product_list)):
                        writer.writerow([new_product_list[i][0],  
                                        new_product_list[i][1]])

            # return a list of new products
            return new_product_list
        
        # if there's no such file, create a new file 
        except FileNotFoundError:
            with open('Izumi_product_list.csv','w') as csvfile:
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
                    if (row[0]==self.output_date) and (row[4]=='008'):
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
                                '人工心肺', 
                                '手術室', 
                                '医療機器管理',
                                5,
                                '泉工医科工業', 
                                '新製品',
                                new_product_list[i][0], 
                                new_product_list[i][1], 
                                1])


# # Run the Function

# In[86]:


if __name__=='__main__':
    year=2020
    month=9
    day=13
    izumi=Izumi(year,month,day)
    izumi.to_csv()


# In[ ]:




