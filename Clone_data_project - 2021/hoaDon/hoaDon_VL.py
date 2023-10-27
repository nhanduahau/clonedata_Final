from asyncio import exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import xlsxwriter
import time
#------------------------------------Hóa Đơn--------------------------------------------
nameOfInputFile = "Clone_data_project/hopDong/results/soHopDong_VL.txt"
nameOfOutputFile ="Clone_data_project - 2021/hoaDon/results/HoaDon_VL.xlsx" 
with open("Clone_data_project - 2021/input_month.txt", 'r') as file:
    # Đọc nội dung từ tệp và chia nó thành danh sách
    cacThang = file.read().splitlines()
#--------------------------------------------------------------------------------
arr_HD= []
f = open(nameOfInputFile)
for line in f:   
    arr_HD.append(line.replace("\n",""))
f.close()
workbook = xlsxwriter.Workbook(nameOfOutputFile)
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Số hợp đồng') 
worksheet.write('B1', 'Mã ĐH') 
worksheet.write('C1', 'Tuyến đọc') 
worksheet.write('D1', 'Tên khách hàng') 
worksheet.write('E1', 'Địa chỉ') 
worksheet.write('F1', 'Chỉ số cũ') 
worksheet.write('G1', 'Chỉ số mới') 
worksheet.write('H1', 'Tiêu thụ') 
worksheet.write('I1', 'Mã ĐT giá') 
worksheet.write('J1', 'Thành tiền') 
worksheet.write('K1', 'Phí ĐTDN') 
worksheet.write('L1', 'Phí BVMT') 
worksheet.write('M1', 'Phí VAT') 
worksheet.write('N1', 'Số tiền') 
worksheet.write('O1', 'Số hóa đơn') 
worksheet.write('P1', 'Seri hóa đơn') 
worksheet.write('Q1', 'Ngày lập HĐ') 
worksheet.write('R1', 'Ngày HĐ') 
worksheet.write('S1', 'Ngày kỳ đầu') 
worksheet.write('T1', 'Ngày kỳ cuối') 
worksheet.write('U1', 'Đã TT') 
worksheet.write('V1', 'Người thu') 
worksheet.write('W1', 'Ghi chú')
worksheet.write('X1', 'Tháng/Năm')

driver = webdriver.Chrome()
driver.get("https://app.citywork.vn/gcrm.aspx")
driver.maximize_window()
ID = driver.find_elements(By.CSS_SELECTOR,'div.col_full:nth-child(1)>input')
ID[0].send_keys("AT_Test")

PASS = driver.find_elements(By.CSS_SELECTOR,'div.col_full:nth-child(2)>input')
PASS[0].send_keys("1234567@A")
BTN = driver.find_elements(By.ID,'ctl00_mainContent_login1_LoginCtrl_Login')
BTN[0].click()
time.sleep(5)

driver.find_element(By.XPATH,'/html/body/div[2]/div[2]/div/div/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div').click()
time.sleep(2)
driver.find_elements(By.CSS_SELECTOR,'li.x-boundlist-item')[3].click()
time.sleep(1)

driver.find_element(By.XPATH,'/html/body/div[2]/div[4]/div/table/tbody/tr[17]/td/div').click()
time.sleep(1)
thang = driver.find_element(By.NAME,'thang')
count = 1

for t in cacThang:
    driver.find_element(By.NAME,'soHopDong').clear()
    thang.clear()
    for c in t:
        thang.send_keys(c)
    time.sleep(2)
    thang.send_keys(Keys.DELETE)
    time.sleep(2)
    maxPage = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
    while(maxPage>900):        
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
        driver.find_element(By.NAME,'soHopDong').clear()
        driver.find_element(By.NAME,'soHopDong').send_keys(i)
        driver.find_element(By.NAME,'soHopDong').send_keys(Keys.ENTER)
        time.sleep(1)
        table = driver.find_elements(By.CSS_SELECTOR,"table.x-grid-table")
        stt = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row")
        print(str(len(stt)) + " - "+ i )
        if(len(stt)>0):
            page = 0
            maxPage = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
            while(page != maxPage):
                table = driver.find_elements(By.CSS_SELECTOR,"table.x-grid-table")
                maHopDong = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(6)")
                for d in range(len(maHopDong)):
                    if(maHopDong[d].text == i):
                        worksheet.write(count,0 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(6)")[d].text)
                        worksheet.write(count,1 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(7)")[d].text)
                        worksheet.write(count,2 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(8)")[d].text)
                        worksheet.write(count,3 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(9)")[d].text)
                        worksheet.write(count,4 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(10)")[d].text)
                        worksheet.write(count,5 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(13)")[d].text)
                        worksheet.write(count,6 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(14)")[d].text)
                        worksheet.write(count,7 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(16)")[d].text)
                        worksheet.write(count,8 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(17)")[d].text)
                        worksheet.write(count,9 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(18)")[d].text)
                        worksheet.write(count,10, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(19)")[d].text)
                        worksheet.write(count,11, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(20)")[d].text)
                        worksheet.write(count,12, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(23)")[d].text)
                        worksheet.write(count,13, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(25)")[d].text)
                        worksheet.write(count,14, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(27)")[d].text)
                        worksheet.write(count,15, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(28)")[d].text)
                        worksheet.write(count,16, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(29)")[d].text)
                        worksheet.write(count,17, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(30)")[d].text)
                        worksheet.write(count,18, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(31)")[d].text)
                        worksheet.write(count,19, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(32)")[d].text)
                        worksheet.write(count,20, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(33)")[d].text)
                        worksheet.write(count,21, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(34)")[d].text)
                        worksheet.write(count,22, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(35)")[d].text)
                        worksheet.write(count,23, t)
                        break
                page = int(driver.find_element(By.NAME,'inputItem').get_attribute('value'))
                driver.find_element(By.CSS_SELECTOR,'span.x-tbar-page-next').click()
                time.sleep(1)
        count+=1
workbook.close()
driver.quit()