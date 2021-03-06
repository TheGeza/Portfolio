#This scaper was created during the course in Linear And Logistic Regression at LTH(Lunds Tekniska Högskola)
#
#Eddi Leino Johansson
#
#pip3 install openpyxl, selenium, ... etc.
import requests
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from geopy.geocoders import Nominatim
import haversine as hs
from haversine import Unit
import time
import math

#Filepath - Where would you like to save your excel sheet with all the data.
filepath = '/Users/eddijohansson/Desktop/Lin.Log/Bostadspriser.xlsx'
workbook = load_workbook(filepath)
#Names all the sheets in the workbook(excel sheet).
sheet = workbook['Data']
sheetGeo = workbook['GeoData']

#Starts the Chrome driver, which can be installed from https://chromedriver.chromium.org/. Make sure to change file path
wd = webdriver.Chrome('/Users/eddijohansson/Desktop/Lin.Log/chromedriver-3')
#Wd.get('url'). The URL are not unique so just put in what settings you would like and copy paste it.
#The link/url has to be put in furtherdown otherwise page 2 will be in Lund(see function loopen()).
wd.get("https://www.hemnet.se/salda/bostader?location_ids%5B%5D=940042&item_types%5B%5D=villa&item_types%5B%5D=radhus&item_types%5B%5D=bostadsratt&sold_age=6m")
time.sleep(3)

#Creates the columns in your sheet, insertData() used therafter. 
def initiateXL():
    sheet['A1'] = 'Objektnummer'
    sheet['B1'] = 'Adress'
    sheet['C1'] = 'Fastighetspriset'
    sheet['D1'] = 'Kvadratmeter'
    sheet['E1'] = 'Tomtarea'
    sheet['F1'] = 'Byggnadsår'
    sheet['G1'] = 'Avgift/mån'
    sheet['H1'] = 'Avstånd till centralen'
    sheet['I1'] = 'Fastighetstyp'
    sheet['J1'] = 'Våning'
    sheet['K1'] = 'Driftskostnad'
    sheet['L1'] = 'Biarea'
    workbook.save('/Users/eddijohansson/Desktop/Lin.Log/Bostadspriser.xlsx')    

#Has to be uncommented to initiate.
# initiateXL()
#Don't change
def getData(index, url, name):
    parent = wd.current_url
    wd.get(url)
    time.sleep(3)
    pris = wd.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/div[1]/div[1]/div[2]/div/span[2]").text
    # print(wd.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/div[1]/div[1]/div[3]/dl[2]/dt[1]").text)
    # print(wd.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/div[1]/div[1]/div[3]/dl[2]/dd[1]").text)
    listdt = (wd.find_elements_by_tag_name('dt'))
    listdd = (wd.find_elements_by_tag_name('dd'))
    i=0
    while i < len(listdt):
        listdt[i] = listdt[i].text
        i += 1
    i=0
    while i < len(listdd):
        listdd[i] = listdd[i].text
        i += 1
    insertData(listdt,listdd,index,name, pris, parent)
#Just change the save location
def insertData(listOrd, listSif, index, gata, salupris, parent):
    place = 'C'+str(index)
    sheet[place] = salupris
    place = 'B'+str(index)
    sheet[place] = gata
    place = 'A' + str(index)
    sheet[place] = str(index-1)
    listMatch = ['Bostadstyp','I','Boarea','D','Våning','J','Byggår','F','Avgift/månad','G','Driftskostnad','K', 'Tomtarea','E','Biarea','L']
    i = 0
    till = 0
    while i < len(listOrd):
        j=0
        while j < len(listMatch):
            if listMatch[j] == listOrd[i]:
                place = listMatch[j+1] + str(index)
                sheet[place] = listSif[i]
            j += 2
        i += 1
    workbook.save('/Users/eddijohansson/Desktop/Lin.Log/Bostadspriser.xlsx')
    time.sleep(1)
    wd.get(parent)
    time.sleep(1)

#This one is used to retrive GeoData for all your data points.
def geoData(address, index):
    pos = 'A' + str(index)
    sheetGeo[pos] = (index-1)
    pos = 'B' + str(index)
    sheetGeo[pos] = address
    try:
        maps = Nominatim(user_agent="my_app_scrapy")
        location = maps.geocode(address +' Lund')
        pos = 'C' + str(index)
        sheetGeo[pos] = str((location.latitude, location.longitude))
        Stadsparken = [(55.69976, 13.18855), (55.70094, 13.18563), (55.69892, 13.18374), (55.69674, 13.18606)]
        Botaniska = [(55.70535, 13.20151),(55.70515, 13.20441),(55.70263, 13.20207)]
        StHans = [(55.72395, 13.187),(55.72403, 13.19277), (55.72145, 13.19398),(55.72138, 13.18621)]
        vectirSp = ['D','E','F','G']
        vectirBo = ['H','I','J']
        vectirSt = ['K','L','M','N']
        i = 0
        while i < len(vectirSp):
            distance = hs.haversine(Stadsparken[i], ((location.latitude, location.longitude)))
            pos = vectirSp[i] + str(index)
            sheetGeo[pos] = distance
            i += 1
        i = 0
        while i < len(vectirBo):
            distance = hs.haversine(Botaniska[i], ((location.latitude, location.longitude)))
            pos = vectirBo[i] + str(index)
            sheetGeo[pos] = distance
            i += 1
        i = 0
        while i < len(vectirSt):
            distance = hs.haversine(StHans[i], ((location.latitude, location.longitude)))
            pos = vectirSt[i] + str(index)
            sheetGeo[pos] = distance
            i += 1
    except:
        pass
    
#The loop which runs the whole program.
def loopen():
    index = 2
    page = 1
    while page < 20:
        a = 2
        while a < 52:
            try:
                test = "/html/body/div[4]/div/div[5]/div[1]/div[3]/ul/li["+str(a)+"]/a"
                test2 = "/html/body/div[4]/div/div[5]/div[1]/div[3]/ul/li["+str(a)+"]/a/div/div[1]/h2/span[2]"
                url = wd.find_element(By.XPATH, test).get_attribute('href')
                name = wd.find_element(By.XPATH,test2).text
            except:
                test = "/html/body/div[4]/div/div[5]/div[1]/div[3]/ul/li["+str(a+1)+"]/a"
                test2 = "/html/body/div[4]/div/div[5]/div[1]/div[3]/ul/li["+str(a+1)+"]/a/div/div[1]/h2/span[2]"
                url = wd.find_element(By.XPATH, test).get_attribute('href')
                name = wd.find_element(By.XPATH,test2).text
                a += 1
            getData(index,url,name)
            geoData(name, index)
            index += 1
            a += 1
        time.sleep(1)
        page += 1
        #Detta url måste bytas till önskad stad, där +str(page)+ ska ersätta page i uralet.
        urle = 'https://www.hemnet.se/salda/bostader?item_types%5B%5D=villa&item_types%5B%5D=radhus&item_types%5B%5D=bostadsratt&location_ids%5B%5D=940042&page='+str(page)+'&sold_age=6m'
        wd.get(urle)


#Not of interse if you didn't miss a few geoData points and missed them and want to add them past tense.
def newGeo():
    maps = Nominatim(user_agent="my_app_scrapy")
    Stortorget =(55.702796187223306, 13.193062425601825)
    Grand_hotel = (55.70389170105121, 13.189021107285864)
    Lunds_stadsbibliotek =  (55.706698157832015, 13.19106954175583)
    Lundagard = (55.70486753537952, 13.193829711080806)
    Saluhallen = (55.7018167226107, 13.195036867923513)
    i=1
    while i < 807:
        p = 'B' + str(i+1)
        add = sheetGeo[p].value +'Lund'
        try:
            location = maps.geocode(add)
        except:
            pass
        if(location != None):
            distance = hs.haversine(Stortorget, ((location.latitude, location.longitude)))
            p = 'O'+str(i+1)
            sheetGeo[p] = distance
            distance = hs.haversine(Grand_hotel, ((location.latitude, location.longitude)))
            p = 'P'+str(i+1)
            sheetGeo[p] = distance
            distance = hs.haversine(Lunds_stadsbibliotek, ((location.latitude, location.longitude)))
            p = 'Q'+str(i+1)
            sheetGeo[p] = distance
            distance = hs.haversine(Lundagard, ((location.latitude, location.longitude)))
            p = 'R'+str(i+1)
            sheetGeo[p] = distance
            distance = hs.haversine(Saluhallen, ((location.latitude, location.longitude)))
            p = 'S'+str(i+1)
            sheetGeo[p] = distance
        i +=1
    print(add)
    workbook.save('/Users/eddijohansson/Desktop/Lin.Log/Bostadspriser.xlsx')

#Clicks away the cookies.
try:
    knapp = wd.find_element(By.XPATH,"/html/body/div[9]/div/div/div/div/div/div[2]/div[2]/div[2]/button")
    knapp.click()
except:
    knapp = wd.find_element(By.XPATH,"/html/body/div[10]/div/div/div/div/div/div[2]/div[2]/div[2]/button")
    knapp.click()
time.sleep(1)

#newGeo()
#Runs the program through loop()
loopen()

