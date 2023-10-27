from asyncio import exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
import mouse
from selenium.webdriver.common.action_chains import ActionChains
import xlsxwriter
import time

start = time.time()
driver = webdriver.Chrome()
driver.get("https://app.citywork.vn/gcrm.aspx")
driver.maximize_window()
#------------------------------------Số Đọc Chỉ Số--------------------------------------------
nameOfOutputFile ="Clone_data_project - 2019/sodoc/results/soDoc_TT.xlsx" 
with open("Clone_data_project - 2019/input_month.txt", 'r') as file:
    # Đọc nội dung từ tệp và chia nó thành danh sách
    cacThang = file.read().splitlines()
#--------------------------------------------------------------------------------
workbook = xlsxwriter.Workbook(nameOfOutputFile)
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Tuyến đọc')
worksheet.write('B1', 'Cán bộ đọc')
worksheet.write('C1', 'Tên sổ')
worksheet.write('D1', 'Chưa ghi')
worksheet.write('E1', 'Chốt sổ')
worksheet.write('F1', 'Trang thái')
worksheet.write('G1', 'Ngày chốt')
worksheet.write('H1', 'Hóa đơn')
worksheet.write('I1', 'Thứ tự')
worksheet.write('J1', 'Tuyến đọc') 
worksheet.write('K1', 'Số HĐ')
worksheet.write('L1', 'Seri ĐH') 
worksheet.write('M1', 'Tên khách hàng')
worksheet.write('N1', 'Địa chỉ')
worksheet.write('O1', 'T.Thu') 
worksheet.write('P1', 'Chỉ số cũ') 
worksheet.write('Q1', 'Chỉ số mới') 
worksheet.write('R1', 'Tiêu thụ') 
worksheet.write('S1', 'Ngày đọc') 
worksheet.write('T1', 'Ngày đồng bộ') 
worksheet.write('U1', 'Trạng thái ghi') 
worksheet.write('V1', 'Ngày Đầu kỳ') 
worksheet.write('W1', 'Ngày cuối kỳ') 
worksheet.write('X1', 'Chỉ số đầu cũ') 
worksheet.write('Y1', 'Chỉ số cuối cũ') 
worksheet.write('Z1', 'Trạng thái ĐH') 
worksheet.write('AA1', 'Ghi chú chỉ số')
worksheet.write('AB1', 'Hình ảnh')  
worksheet.write('AC1', 'Tháng/Năm')  
ID = driver.find_elements(By.CSS_SELECTOR,'div.col_full:nth-child(1)>input')
ID[0].send_keys("AT_Test")
time.sleep(1)
PASS = driver.find_elements(By.CSS_SELECTOR,'div.col_full:nth-child(2)>input')
PASS[0].send_keys("1234567@A")
BTN = driver.find_elements(By.ID,'ctl00_mainContent_login1_LoginCtrl_Login')
BTN[0].click()
time.sleep(5)

driver.find_element(By.ID,'ext-gen1172').click()
time.sleep(2)
driver.find_elements(By.CSS_SELECTOR,'li.x-boundlist-item')[2].click()
time.sleep(1)




time.sleep(1)
driver.find_element(By.ID,'ext-gen1212').click()
time.sleep(3)
thang = driver.find_element(By.NAME,'thang')
count = 1
test = []
count_soDoc =1
for t in cacThang:
    thang.clear()
    for c in t:
        thang.send_keys(c)
    time.sleep(2)
    thang.send_keys(Keys.DELETE)
    time.sleep(2)
    maxPage = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
    while(maxPage>4000):        
        thang.clear()
        thang.send_keys("08/2023")
        thang.clear()
        for c in t:
            thang.send_keys(c)
        time.sleep(2)
        thang.send_keys(Keys.DELETE)
        time.sleep(2)
        maxPage = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
    page = 0
    while(page != maxPage):
        table = driver.find_elements(By.CSS_SELECTOR,"table.x-grid-table")
        tr = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row")
        tuyenDoc = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(3)")
        canBoDoc = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(4)")
        tenSo = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(5)")
        chuaGhi = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(6)")
        chotSo = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(7)")
        trangThai = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(8)")
        ngayChot = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(9)")
        hoaDon = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(11)")
        quanLyHopDong = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(12)")
        for index in range(len(tr)):
            arr = []
            arr.append(tuyenDoc[index].text)
            arr.append(canBoDoc[index].text)
            arr.append(tenSo[index].text)
            arr.append(chuaGhi[index].text)
            arr.append(chotSo[index].text)
            arr.append(trangThai[index].text)
            arr.append(ngayChot[index].text)
            arr.append(hoaDon[index].text)
            quanLyHopDong[index].find_element(By.CSS_SELECTOR,"img.iconInvoiceRed").click()
            time.sleep(2)
#-------------------------------------------
            table = driver.find_elements(By.CSS_SELECTOR,"table.x-grid-table")
            stt = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row")
            page_c = 0
            maxPage_c = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[4].text.replace("của ",""))
            ii = 0
            while(page_c != maxPage_c):
                table = driver.find_elements(By.CSS_SELECTOR,"table.x-grid-table")
                row = table[4].find_elements(By.CSS_SELECTOR,"tr.x-grid-row")
                for d in row:
                    d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[1].find_element(By.CSS_SELECTOR,"img.x-action-col-1").click()
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
                    worksheet.write(count, 0, arr[0])
                    worksheet.write(count, 1, arr[1])
                    worksheet.write(count, 2, arr[2])
                    worksheet.write(count, 3, arr[3])
                    worksheet.write(count, 4, arr[4])
                    worksheet.write(count, 5, arr[5])
                    worksheet.write(count, 6, arr[6])
                    worksheet.write(count, 7, arr[7])
                    worksheet.write(count,8 , d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[2 ].text)
                    worksheet.write(count,9 , d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[3 ].text)
                    worksheet.write(count,10, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[4 ].text)
                    worksheet.write(count,11, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[6 ].text)
                    worksheet.write(count,12, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[6 ].text)
                    worksheet.write(count,13, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[9].text)
                    worksheet.write(count,14, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[10].text)
                    worksheet.write(count,15, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[12].text)
                    worksheet.write(count,16, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[13].text)
                    worksheet.write(count,17, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[15].text)
                    worksheet.write(count,18, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[18].text)
                    worksheet.write(count,19, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[19].text)
                    worksheet.write(count,20, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[20].text)
                    worksheet.write(count,21, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[21].text)
                    worksheet.write(count,22, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[22].text)
                    worksheet.write(count,23, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[23].text)
                    worksheet.write(count,24, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[24].text)
                    worksheet.write(count,25, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[26].text)
                    worksheet.write(count,26, d.find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[27].text)
                    worksheet.write(count,27, src_image)
                    worksheet.write(count,28, t)
                    count+=1
                driver.find_elements(By.CSS_SELECTOR,'span.x-tbar-page-next')[1].click()
                page_c = int(driver.find_elements(By.NAME,'inputItem')[1].get_attribute('value'))
                time.sleep(1)
#-------------------------------------------            
            driver.find_elements(By.CSS_SELECTOR,"a.x-tab-close-btn")[1].click()
            count_soDoc +=1
        page = int(driver.find_element(By.NAME,'inputItem').get_attribute('value'))
        driver.find_element(By.CSS_SELECTOR,'span.x-tbar-page-next').click()
        time.sleep(2)
        print('---------------Trang '+str(page)+' hoàn thành---------------')
    print("---------------Tháng "+t + " hoàn thành---------------")
workbook.close()
driver.quit()
end = time.time()
print(str(end - start)+"s")