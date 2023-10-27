from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time
from selenium.webdriver.common.keys import Keys
import xlsxwriter
start = time.time()
driver = webdriver.Chrome()
driver.get("https://app.citywork.vn/gcrm.aspx")
driver.maximize_window()
cacThang =["01/2021","02/2021","03/2021","04/2021","05/2021","06/2021","07/2021","08/2021","09/2021","10/2021","11/2021","12/2021"]

workbook = xlsxwriter.Workbook("nhapChiSo_2021.xlsx")
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'STT')
worksheet.write('B1', 'Tuyến đọc')
worksheet.write('C1', 'Số HĐ')
worksheet.write('D1', 'Seri hợp đồng')
worksheet.write('E1', 'Tên KH')
worksheet.write('F1', 'Địa chỉ')
worksheet.write('G1', 'T.Thu')
worksheet.write('H1', 'Chỉ số cũ')
worksheet.write('I1', 'CHỉ số mới')
worksheet.write('J1', 'Tiêu thụ')
worksheet.write('K1', 'Ngày đọc')
worksheet.write('L1', 'Ngày đồng bộ')
worksheet.write('M1', 'Trạng thái ghi')
worksheet.write('N1', 'Ngày kỳ đầu')
worksheet.write('O1', 'Ngày kỳ cuối')
worksheet.write('P1', 'Chỉ số đầu cũ')
worksheet.write('Q1', 'Chỉ số cuối cũ')
worksheet.write('R1', 'Trạng thái ĐH')
worksheet.write('S1', 'Ghi chú')
worksheet.write('T1', 'Hình ảnh')
worksheet.write('U1', 'Thông tin')

ID = driver.find_elements(By.CSS_SELECTOR,'div.col_full:nth-child(1)>input')
ID[0].send_keys("AT_Test")
time.sleep(1)
PASS = driver.find_elements(By.CSS_SELECTOR,'div.col_full:nth-child(2)>input')
PASS[0].send_keys("1234567890Aa")
BTN = driver.find_elements(By.ID,'ctl00_mainContent_login1_LoginCtrl_Login')
BTN[0].click()
driver.find_element(By.ID,'ext-gen1171').click()
time.sleep(5)
driver.find_elements(By.CSS_SELECTOR,'li.x-boundlist-item')[4].click()
time.sleep(1)
driver.find_element(By.ID,'ext-gen1213').click()
time.sleep(2)
thang = driver.find_element(By.NAME,'thang')
count = 1

for i in cacThang:
    time.sleep(1)
    thang.clear()
    for c in i:
        thang.send_keys(c)
    time.sleep(5)
    maxPage = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
    while(maxPage>800):        
        thang.clear()
        thang.send_keys("08/2023")
        thang.clear()
        for c in i:
            thang.send_keys(c)
        time.sleep(5)
        maxPage = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
    print(i+ " có "+str(maxPage)+" trang")
    page = 1
    input("ok")
    while(page != maxPage):
        table = driver.find_elements(By.CSS_SELECTOR,"table.x-grid-table")
        row = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row")
        for index in range(len(row)):
            row[index].find_element(By.CSS_SELECTOR,"img.x-action-col-1").click()
            time.sleep(2)
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
            chu_thich = str(row[index].find_element(By.CSS_SELECTOR,"img.x-action-col-2").get_attribute("data-qtip")).replace("<b>"," ").replace("</b>","")
            worksheet.write(count, 0, row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[2].text)
            worksheet.write(count, 1, row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[3].text)
            worksheet.write(count, 2, row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[4].text)
            worksheet.write(count, 3, row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[6].text)
            worksheet.write(count, 4, row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[8].text)
            worksheet.write(count, 5, row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[9].text)
            worksheet.write(count, 6, row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[10].text)
            worksheet.write(count, 7, row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[11].text)
            worksheet.write(count, 8, row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[12].text)
            worksheet.write(count, 9, row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[14].text)
            worksheet.write(count, 10,row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[17].text)
            worksheet.write(count, 11,row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[18].text)
            worksheet.write(count, 12,row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[19].text)
            worksheet.write(count, 13,row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[20].text)
            worksheet.write(count, 14,row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[21].text)
            worksheet.write(count, 15,row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[22].text)
            worksheet.write(count, 16,row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[23].text)
            worksheet.write(count, 17,row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[24].text)
            worksheet.write(count, 18,row[index].find_elements(By.CSS_SELECTOR,"td.x-grid-cell")[25].text)
            worksheet.write(count, 19,src_image)
            worksheet.write(count, 20,chu_thich)
            print("---- "+str(count)+" ----")
            count +=1
        page = int(driver.find_element(By.NAME,'inputItem').get_attribute('value'))
        driver.find_element(By.CSS_SELECTOR,'span.x-tbar-page-next').click()
        time.sleep(2)
        print('---------------Trang '+str(page)+' hoàn thành---------------')
    print("---------------Tháng "+i+ " hoàn thành---------------")
workbook.close()
driver.quit()
end = time.time()
print(str(end - start)+"s")