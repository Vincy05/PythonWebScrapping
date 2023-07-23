from lib2to3.pgen2.driver import Driver
from sqlite3 import Cursor
from bs4 import BeautifulSoup
import requests
import pandas as pd
import pyodbc as pc
import sys 
import pypyodbc 
import sqlite3
import sqlalchemy 
from sqlalchemy import create_engine
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException
from sqlalchemy.sql import text
from sqlalchemy.orm import sessionmaker
import openpyxl



pageno=[]
price_list = []
bed_list = []
address_list = []
tag_list = []
Property_Tax =[]
Last_Updated=[]
Date_Listed=[]
PropertyType=[]
Subdivision =[]
SquareFootage=[]
FullBathrooms=[]
HalfBathrooms=[]
Heating=[]
Stories=[]
PetPolicy =[]
Basement=[]
Area=[]
Flooring=[]
Roof=[]
Sewer=[]
ParkingFeatures=[]
FirePlaceFeatures=[]
ExteriorFeatures=[]
LastPageNo=[]
ZipCode=[]

options = Options()
options.add_argument("--remote-debugging-port=9230")
options.add_experimental_option('excludeSwitches', ['enable-logging'])


service = Service(executable_path=r"C:\Users\Vincy\Downloads\chromedriver_win32\chromedriver.exe") 
driver = webdriver.Chrome(service=service) 
url ="https://www.remax.ca/find-real-estate?address=British+Columbia%2C+Canada&pageNumber=1"
driver.get(url) 
time.sleep(10)
soup = BeautifulSoup(driver.page_source,'html.parser')
mx =   soup.findAll('a', class_='MuiButtonBase-root MuiButton-root rootSizeIcon remax-button_square___st8Q commercialOutlined page-control_pageButton__dUo1o residentialDarkColour MuiButton-text remax-button_buttonText__saWK7 remax-button_buttonTextIcon__NQa87')
for i in mx:
    mxnum=i.get_text( )
    mxnum1=3 #determining the number of pages to be extracted from web page
    LastPageNo.append(mxnum1)
print(LastPageNo[-1])
for page in range(1,int(LastPageNo[-1])):
    url1="https://www.remax.ca/find-real-estate?address=British+Columbia%2C+Canada&pageNumber="
    url2 = url1+str(page)+""
    driver.get(url2)
    time.sleep(10)
    soup = BeautifulSoup(driver.page_source,'html.parser')
    listings = soup.find_all('div', class_='listing-card_root__UG576 search-gallery_galleryCardRoot__7HbLb')
    for i in listings: 
        try:
            price = i.find('h2', class_='listing-card_price__sL9TT').get_text( )
        except Exception as e:
            price.append('null') 
        price_list.append(price)

        try:
            bed = i.find('div', class_='property-details_detailsRow__nilLP').get_text( )
        except Exception as e:
            bed.append('null')
        bed_list.append(bed)

        try:
            address = i.find('div', class_='listing-address_root__PP_Ky listing-card_address___bLLz').get_text( )
        except Exception as e:
            address.append('null')   
        address_list.append(address)     
        

        try:
            tag = i.find('div', class_='listing-tag-container_listingTags__z_AhT').get_text( )
        except Exception as e:
            tag.append('null') 
        tag_list.append(tag)

    for i in listings:
        NewTab = i.find('a', class_='listing-card_listingCard__G6M8g').get('href')
        html_text1 = requests.get(NewTab)
        soup1 = BeautifulSoup(html_text1.content, 'html.parser')
        hme = soup1.find_all('li',  class_='bullet-section_bulletPointRow__4pBp6')
        zip = soup1.find_all('span',  class_='listing-summary_cityLine__YxXgL listing-address_splitLines__pLZIy')
        Property =[]  

        for x in hme:
            main=(x.get_text( ))

            if 'Property Tax' in main:
                hme1 =  main.split(":")[1].strip() 
                Property_Tax.append(hme1)
                break

        else:
            Property_Tax.append("Null")  

        for x in hme:
            main=(x.get_text( ))

            if 'Last Updated' in main:
                hme2 =  main.split(":")[1].strip() 
                Last_Updated.append(hme2)
                break

        else:
            Last_Updated.append("Null")  

        for x in hme:
            main=(x.get_text( ))

            if 'Date Listed' in main:
                hme3 =  main.split(":")[1].strip() 
                Date_Listed.append(hme3)
                break

        else:
            Date_Listed.append("Null")  

        for x in hme:
            main=(x.get_text( ))

            if 'Property Type' in main:
                hme4 =  main.split(":")[1].strip() 
                PropertyType.append(hme4)
                break

        else:
            PropertyType.append("Null")  

        for x in hme:
            main=(x.get_text( ))

            if 'Subdivision' in main:
                hme5 =  main.split(":")[1].strip() 
                Subdivision.append(hme5)
                break

        else:
            Subdivision.append("Null")  

        for x in hme:
            main=(x.get_text( ))

            if 'Square Footage' in main:
                hme6 =  main.split(":")[1].strip() 
                SquareFootage.append(hme6)
                break

        else:
            SquareFootage.append("Null")  

        for x in hme:
            main=(x.get_text( ))

            if 'Full Bathrooms' in main:
                hme7 =  main.split(":")[1].strip() 
                FullBathrooms.append(hme7)
                break

        else:
            FullBathrooms.append("Null")  

        for x in hme:
            main=(x.get_text( ))

            if 'Half Bathrooms' in main:
                hme8 =  main.split(":")[1].strip() 
                HalfBathrooms.append(hme8)
                break

        else:
            HalfBathrooms.append("Null")  


        for x in hme:
            main=(x.get_text( ))

            if 'Heating' in main:
                hme9 =  main.split(":")[1].strip() 
                Heating.append(hme9)
                break

        else:
            Heating.append("Null")  

        for x in hme:
            main=(x.get_text( ))

            if 'Flooring' in main:
                hme10 =  main.split(":")[1].strip() 
                Flooring.append(hme10)
                break

        else:
            Flooring.append("Null")  


        for x in hme:
            main=(x.get_text( ))

            if 'Basement' in main:
                hme11 =  main.split(":")[1].strip() 
                Basement.append(hme11)
                break

        else:
            Basement.append("Null")  


        for x in hme:
            main=(x.get_text( ))

            if 'Roof' in main:
                hme12 =  main.split(":")[1].strip() 
                Roof.append(hme12)
                break

        else:
            Roof.append("Null")  

        for x in hme:
            main=(x.get_text( ))

            if 'Parking Features' in main:
                hme13=  main.split(":")[1].strip() 
                ParkingFeatures.append(hme13)
                break

        else:
            ParkingFeatures.append("Null")  

        for x in hme:
            main=(x.get_text( ))

            if 'Fireplace Features' in main:
                hme14 =  main.split(":")[1].strip() 
                FirePlaceFeatures.append(hme14)
                break

        else:
            FirePlaceFeatures.append("Null")  

        for x in hme:
            main=(x.get_text( ))

            if 'Exterior Features' in main:
                hme15 =  main.split(":")[1].strip() 
                ExteriorFeatures.append(hme15)
                break

        else:
            ExteriorFeatures.append("Null")  


        for c in zip:
            zip1 =  c.get_text( ) 

        try:
            ZipCode.append(zip1)
        except Exception as e:
            ZipCode.append('null') 

df =pd.DataFrame({'Price': price_list,'BedType': bed_list, 'Address_1' : address_list , 'Tag' : tag_list ,  'Property_Tax_Price': Property_Tax
, 'Last_Updated':Last_Updated, 'Date_Listed':Date_Listed, 'PropertyType':PropertyType , 'SquareFootage' :SquareFootage
, 'Subdivision':Subdivision,'FullBathrooms':FullBathrooms, 'HalfBathrooms' : HalfBathrooms,'Heating': Heating , 'Flooring': Flooring
, 'Basement': Basement ,'Roof':Roof , 'FirePlaceFeatures': FirePlaceFeatures,'ParkingFeatures': ParkingFeatures
, 'ExteriorFeatures': ExteriorFeatures, 'ZipCode': ZipCode })

# print(df)
# prints the complete  housing data prices 

#To write the code to excel 
df.to_excel("HousingPrice.xlsx")


# DB CONNECTION
server = "LAPTOP-DLC1JV7R"
database = "AdventureWorks2019"
driver = "ODBC+Driver+17+for+SQL+Server"
url = f"mssql+pyodbc://{server}/{database}?trusted_connection=yes&driver={driver}"
engine = sqlalchemy.create_engine(url)

Session = sessionmaker(bind=engine)
session = Session()
session.execute(text('''TRUNCATE TABLE Housing_test2'''))
session.commit()
session.close()

df.to_sql('Housing_test2', engine, if_exists='append', index = False, )



