from asyncio import exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import xlsxwriter
import time
#------------------------------------Nhập Chỉ Số--------------------------------------------
nameOfInputFile = "dealine/soHopDong_VL.txt"
nameOfOutputFile ="dealine/ketQua/nhapChiSo_VL.xlsx" 
cacThang =["07/2023","08/2023"]
#--------------------------------------------------------------------------------
arr_HD= []
f = open(nameOfInputFile)
for line in f:   
    arr_HD.append(line.replace("\n",""))
f.close()
workbook = xlsxwriter.Workbook(nameOfOutputFile)
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Thứ tự')
worksheet.write('B1', 'Tuyến đọc') 
worksheet.write('C1', 'Số HĐ')
worksheet.write('D1', 'Seri ĐH') 
worksheet.write('E1', 'Tên khách hàng')
worksheet.write('F1', 'Địa chỉ')
worksheet.write('G1', 'T.Thu') 
worksheet.write('H1', 'Chỉ số cũ') 
worksheet.write('I1', 'Chỉ số mới') 
worksheet.write('J1', 'Tiêu thụ') 
worksheet.write('K1', 'Ngày đọc') 
worksheet.write('L1', 'Ngày đồng bộ') 
worksheet.write('M1', 'Trạng thái ghi') 
worksheet.write('N1', 'Ngày Đầu kỳ') 
worksheet.write('O1', 'Ngày cuối kỳ') 
worksheet.write('P1', 'Chỉ số đầu cũ') 
worksheet.write('Q1', 'Chỉ số cuối cũ') 
worksheet.write('R1', 'Trạng thái ĐH') 
worksheet.write('S1', 'Ghi chú chỉ số')
worksheet.write('T1', 'Hình ảnh')  
worksheet.write('U1', 'Tháng/Năm')
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

driver.find_element(By.ID,'ext-gen1172').click()
time.sleep(2)
driver.find_elements(By.CSS_SELECTOR,'li.x-boundlist-item')[3].click()
time.sleep(1)

driver.find_element(By.ID,'ext-gen1214').click()
time.sleep(1)
thang = driver.find_element(By.NAME,'thang')
count = 1

for t in cacThang:
    thang.clear()
    for c in t:
        thang.send_keys(c)
    time.sleep(2)
    thang.send_keys(Keys.DELETE)
    time.sleep(2)
    maxPage = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
    while(maxPage>1000):        
        thang.clear()
        thang.send_keys("08/2023")
        thang.clear()
        for c in t:
            thang.send_keys(c)
        time.sleep(2)
        thang.send_keys(Keys.DELETE)
        time.sleep(2)
        maxPage = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
    for i in arr_HD:
        driver.find_element(By.NAME,'txtMaDongHo').clear()
        driver.find_element(By.NAME,'txtMaDongHo').send_keys(i)
        driver.find_elements(By.CSS_SELECTOR,'button.x-btn-center')[1].click()
        time.sleep(0.5)
        table = driver.find_elements(By.CSS_SELECTOR,"table.x-grid-table")
        stt = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row")
        print(str(len(stt)) + " - "+ i )
        if(len(stt)>0):
            page = 0
            maxPage = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
            while(page != maxPage):
                table = driver.find_elements(By.CSS_SELECTOR,"table.x-grid-table")
                maHopDong = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(5)")
                for d in range(len(maHopDong)):
                    if(maHopDong[d].text == i):
                        table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(2)")[d].find_element(By.CSS_SELECTOR,"img.x-action-col-1").click()
                        time.sleep(1)
                        popup = driver.find_element(By.ID,"wnShowChoseItem")
                        src_image =""
                        try:
                            image = popup.find_element(By.ID,"tab-fviewimage-body")
                            image = image.find_element(By.CSS_SELECTOR,"img.x-component")
                            nameOfImage = popup.find_elements(By.CSS_SELECTOR,"table.x-grid-table")[1].find_elements(By.CSS_SELECTOR,"tr.x-grid-row")
                            for name in nameOfImage:
                                name.click()
                                time.sleep(1)
                                src_image += image.get_attribute("src") + " "
                        except NoSuchElementException:
                            pass
                        popup.find_element(By.CSS_SELECTOR,"img.x-tool-close").click()
                        worksheet.write(count,0 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(3)")[d].text)
                        worksheet.write(count,1 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(4)")[d].text)
                        worksheet.write(count,2 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(5)")[d].text)
                        worksheet.write(count,3 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(7)")[d].text)
                        worksheet.write(count,4 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(9)")[d].text)
                        worksheet.write(count,5 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(10)")[d].text)
                        worksheet.write(count,6 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(11)")[d].text)
                        worksheet.write(count,7 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(13)")[d].text)
                        worksheet.write(count,8 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(14)")[d].text)
                        worksheet.write(count,9 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(16)")[d].text)
                        worksheet.write(count,10, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(19)")[d].text)
                        worksheet.write(count,11, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(20)")[d].text)
                        worksheet.write(count,12, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(21)")[d].text)
                        worksheet.write(count,13, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(22)")[d].text)
                        worksheet.write(count,14, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(23)")[d].text)
                        worksheet.write(count,15, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(24)")[d].text)
                        worksheet.write(count,16, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(25)")[d].text)
                        worksheet.write(count,17, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(27)")[d].text)
                        worksheet.write(count,18, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(28)")[d].text)
                        worksheet.write(count,19, src_image)
                        worksheet.write(count,20, t)
                        break
                page = int(driver.find_element(By.NAME,'inputItem').get_attribute('value'))
                driver.find_element(By.CSS_SELECTOR,'span.x-tbar-page-next').click()
                time.sleep(1)
        count+=1
workbook.close()
driver.quit()