from asyncio import exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import xlsxwriter
import time
#------------------------------------Thanh Toán--------------------------------------------
nameOfInputFile = "Clone_data_project/hopDong/results/soHopDong_TT.txt"
nameOfOutputFile ="Clone_data_project - 2022/thanhtoan/results/thanhToan_TT.xlsx" 
with open("Clone_data_project - 2022/input_month.txt", 'r') as file:
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
worksheet.write('A1', 'Thanh toán') 
worksheet.write('B1', 'Ngày thu') 
worksheet.write('C1', 'Người thu tiền') 
worksheet.write('D1', 'Tổng tiền') 
worksheet.write('E1', 'Tổng tiền đã TT')
worksheet.write('F1', 'Số hợp đồng')
worksheet.write('G1', 'Tuyến đọc')
worksheet.write('H1', 'Tên khách hàng')
worksheet.write('I1', 'Địa chỉ')
worksheet.write('J1', 'Điện thoại') 
worksheet.write('K1', 'Hình thức thanh toán') 
worksheet.write('L1', 'Seri hóa đơn') 
worksheet.write('M1', 'Số hóa đơn') 
worksheet.write('N1', 'Tiêu thụ') 
worksheet.write('O1', 'Thành tiền') 
worksheet.write('P1', 'Phí ĐTDN') 
worksheet.write('Q1', 'Phí BVMT') 
worksheet.write('R1', 'Phí VAT') 
worksheet.write('S1', 'Ngày lập HĐ') 
worksheet.write('T1', 'Ghi chú thanh toán')
worksheet.write('U1', 'Tháng/Năm')

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
driver.find_elements(By.CSS_SELECTOR,'li.x-boundlist-item')[2].click()
time.sleep(1)
driver.find_element(By.XPATH,'/html/body/div[2]/div[4]/div/table/tbody/tr[29]/td/div').click()

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
        driver.find_element(By.NAME,'txtMaHopDong').clear()
        driver.find_element(By.NAME,'txtMaHopDong').send_keys(i)
        driver.find_element(By.NAME,'txtMaHopDong').send_keys(Keys.ENTER)
        time.sleep(1)
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
                        worksheet.write(count,5 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(11)")[d].text)
                        worksheet.write(count,6 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(12)")[d].text)
                        worksheet.write(count,7 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(13)")[d].text)
                        worksheet.write(count,8 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(14)")[d].text)
                        worksheet.write(count,9 , table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(15)")[d].text)
                        worksheet.write(count,10, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(16)")[d].text)
                        worksheet.write(count,11, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(18)")[d].text)
                        worksheet.write(count,12, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(19)")[d].text)
                        worksheet.write(count,13, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(22)")[d].text)
                        worksheet.write(count,14, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(23)")[d].text)
                        worksheet.write(count,15, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(24)")[d].text)
                        worksheet.write(count,16, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(25)")[d].text)
                        worksheet.write(count,17, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(28)")[d].text)
                        worksheet.write(count,18, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(29)")[d].text)
                        worksheet.write(count,19, table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(30)")[d].text)
                        worksheet.write(count,20, t)
                        break
                page = int(driver.find_element(By.NAME,'inputItem').get_attribute('value'))
                driver.find_element(By.CSS_SELECTOR,'span.x-tbar-page-next').click()
                time.sleep(1)
        count+=1
workbook.close()
driver.quit()