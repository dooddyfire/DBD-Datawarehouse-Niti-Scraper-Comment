import datetime
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
#Fix
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
#Insert file name


#def scrape_url(idx): 
    #pass 
file_name = input("ใส่ชื่อไฟล์เลขทะเบียน นามสกุลด้วย : ")
data = pd.read_excel(file_name)
r = [i for i in data["เลขทะเบียน"]]

#Get bot selenium make sure you can access google chrome
driver = webdriver.Chrome(ChromeDriverManager().install())
driver.get("https://datawarehouse.dbd.go.th/")

#Make sure you have already eliminate recapcha
x = input("Already input recapcha? :")

count = 0
niti_place_lis = []
year_lis = []
executor_lis = []
purpose_money_lis = []
comp_name_lis = []
addr_lis = []
purpose_start_lis = []
web_lis = []
email_lis = []
phone_lis = []
for i in r:
    p = []
    i = "0"+str(i)

    count += 1



    
    driver.get("https://datawarehouse.dbd.go.th/company/profile/"+i[3]+"/"+i)
    print("https://datawarehouse.dbd.go.th/company/profile/"+i[3]+"/"+i)

    soup = BeautifulSoup(driver.page_source,'html.parser')
    comment = soup.find_all('td',{'class':'title'})
    print(comment)
    niti_place = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/table/tbody/tr[3]/td[2]").text 
    niti_place_lis.append(niti_place)
    print(niti_place)

    year = [ c.text for c in driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/table/tbody/tr[7]/td[2]").find_elements(By.CSS_SELECTOR,'a')]
    year_lis.append(" ".join(year))
    print(year)

    executor = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[3]/div[1]/div").text 
    executor_lis.append(executor)
    print(executor)


    purpose_money = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[3]/div[5]/div/p").text 
    purpose_money_lis.append(purpose_money)
    print(purpose_money)

    comp_name = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/div[2]/div[1]/div[1]/h2").text 
    print(comp_name.split(":")[1].strip())
    comp_name_lis.append(comp_name.split(":")[1].strip())

    addr = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[2]/td").text 
    print(addr)
    addr_lis.append(addr)

    purpose_start = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[3]/div[2]/div").text 
    purpose_start_lis.append(purpose_start)
    print(purpose_start)

    web = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[3]/td[2]").text 
    print(web)
    web_lis.append(web)

    tel_idx = driver.page_source.find("โทรศัพท์")
    tel_comment = driver.page_source[tel_idx:tel_idx+70]
    print(tel_comment.split("<td>")[1].replace("</td>","").strip())
    phone_lis.append(tel_comment.split("<td>")[1].replace("</td>","").strip()[:12])

    email_idx = driver.page_source.find("E-mail address")
    email_comment = driver.page_source[email_idx:email_idx+100]
    
    email_raw = email_comment.split("<td>")[1].replace("</td>","").strip()
    email_lis.append(email_raw[0:email_raw.find("-")])
    print(email_comment.split("<td>")[1].replace("</td>","").strip())

    #email = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/table/tbody/comment()[11]")
    #print(email)
    #email_lis.append(email)

    #phone = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/table/tbody/comment()[7]")
    #print(phone)
    #phone_lis.append(phone)

df = pd.DataFrame()
df['เลขทะเบียน'] = r 
df['สถานะนิติบุคคล'] = niti_place_lis
df['รายชื่อกรรมการ'] = executor_lis 
df['วัตถุประสงค์ที่ส่งงบการเงินปีล่าสุด'] = purpose_money_lis 
df['ปีที่ส่งงบการเงิน'] = year_lis 
df['ชื่อบริษัท'] = comp_name_lis 
df['ที่ตั้ง'] = addr_lis 
df['วัตถุประสงค์'] = purpose_start_lis 
df['เว็บไซต์'] = web_lis 
df['Email'] = email_lis 
df['Phone'] = phone_lis 

df.to_excel("result.xlsx")