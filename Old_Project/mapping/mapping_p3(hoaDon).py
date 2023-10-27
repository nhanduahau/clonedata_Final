import pandas as pd

xl = pd.ExcelFile('mapping-p1.xlsx')

df = pd.read_excel(xl, 0, header=None)
max_rows = len(df.iloc[:, 0])
arr_HD= []
for i in range(1,int(max_rows)):
    arr_HD.append(df.at[i, 13])
from asyncio import exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import xlsxwriter
import time

workbook = xlsxwriter.Workbook("mapping-p3.xlsx")
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Thanh toán') 
worksheet.write('B1', 'Ngày thu') 
worksheet.write('C1', 'Người thu tiền') 
worksheet.write('D1', 'Tổng tiền') 
worksheet.write('E1', 'Tổng tiền đã thanh toán') 
worksheet.write('F1', 'Điện thoại') 
worksheet.write('G1', 'Hình thức thanh toán') 
worksheet.write('H1', 'Seri hóa đơn') 
worksheet.write('I1', 'Số hóa đơn') 
worksheet.write('J1', 'Tiêu thụ') 
worksheet.write('K1', 'Thành tiền') 
worksheet.write('L1', 'Phí ĐTDN') 
worksheet.write('M1', 'Phí BVMT') 
worksheet.write('N1', 'Phí VAT') 
worksheet.write('O1', 'Ngày lập HĐ') 
worksheet.write('P1', 'Ghi chú thanh toán') 

driver = webdriver.Chrome()
driver.get("https://app.citywork.vn/gcrm.aspx")
driver.maximize_window()
ID = driver.find_elements(By.CSS_SELECTOR,'div.col_full:nth-child(1)>input')
ID[0].send_keys("AT_Test")

PASS = driver.find_elements(By.CSS_SELECTOR,'div.col_full:nth-child(2)>input')
PASS[0].send_keys("1234567890Aa")
BTN = driver.find_elements(By.ID,'ctl00_mainContent_login1_LoginCtrl_Login')
BTN[0].click()
time.sleep(5)

driver.find_element(By.ID,'ext-gen1240').click()
cacThang =["05/2023","06/2023","07/2023"]
time.sleep(1)
thang = driver.find_element(By.NAME,'thang')
count = 1

for i in cacThang:
    thang.clear()
    for c in i:
        thang.send_keys(c)
    time.sleep(2)
    thang.send_keys(Keys.DELETE)
    time.sleep(2)
    maxPage = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
    while(maxPage>900):        
        thang.clear()
        thang.send_keys("08/2023")
        thang.clear()
        for c in i:
            thang.send_keys(c)
        time.sleep(2)
        thang.send_keys(Keys.DELETE)
        time.sleep(2)
        maxPage = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
    for i in arr_HD:
        driver.find_element(By.NAME,'txtMaHopDong').clear()
        driver.find_element(By.NAME,'txtMaHopDong').send_keys(i)
        driver.find_element(By.NAME,'txtMaHopDong').send_keys(Keys.ENTER)
        time.sleep(0.5)
        table = driver.find_elements(By.CSS_SELECTOR,"table.x-grid-table")
        stt = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row")
        print(str(len(stt)) + " - "+ i )
        if(len(stt)>0):
            page = 0
            maxPage = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
            while(page != maxPage):
                table = driver.find_elements(By.CSS_SELECTOR,"table.x-grid-table")
                maHopDong = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(11)")
                for d in range(len(maHopDong)):
                    if(maHopDong[d].text == i):
                        worksheet.write(count,0, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(3)")[d].text)
                        worksheet.write(count,1, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(4)")[d].text)
                        worksheet.write(count,2, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(5)")[d].text)
                        worksheet.write(count,3, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(9)")[d].text)
                        worksheet.write(count,4, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(10)")[d].text)
                        worksheet.write(count,5, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(15)")[d].text)
                        worksheet.write(count,6, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(16)")[d].text)
                        worksheet.write(count,7, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(18)")[d].text)
                        worksheet.write(count,8, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(19)")[d].text)
                        worksheet.write(count,9, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(22)")[d].text)
                        worksheet.write(count,10, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(23)")[d].text)
                        worksheet.write(count,11, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(24)")[d].text)
                        worksheet.write(count,12, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(25)")[d].text)
                        worksheet.write(count,13, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(28)")[d].text)
                        worksheet.write(count,14, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(29)")[d].text)
                        worksheet.write(count,15, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(30)")[d].text)
                        
                        break
                page = int(driver.find_element(By.NAME,'inputItem').get_attribute('value'))
                driver.find_element(By.CSS_SELECTOR,'span.x-tbar-page-next').click()
                time.sleep(1)
        count+=1
workbook.close()
driver.quit()