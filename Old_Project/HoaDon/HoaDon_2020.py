from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time
from selenium.webdriver.common.keys import Keys
import xlsxwriter
start = time.time()
driver = webdriver.Chrome()
driver.get("https://app.citywork.vn/Secure/Login.aspx")
driver.maximize_window()
cacThang =["01/2020","02/2020","03/2020","04/2020","05/2020","06/2020","07/2020","08/2020","09/2020","10/2020","11/2020","12/2020"]

workbook = xlsxwriter.Workbook("hoaDon2020.xlsx")
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'STT')
worksheet.write('B1', 'Số hợp đồng')
worksheet.write('C1', 'Mã HĐ')
worksheet.write('D1', 'Tuyến dọc')
worksheet.write('E1', 'Tên KH')
worksheet.write('F1', 'Địa chỉ')
worksheet.write('G1', 'Chỉ số cũ')
worksheet.write('H1', 'CHỉ số mới')
worksheet.write('I1', 'Tiêu thụ')
worksheet.write('J1', 'Mã ĐT giá')
worksheet.write('K1', 'Thành tiền')
worksheet.write('L1', 'Phí DTĐN')
worksheet.write('M1', 'Phí BVMT')
worksheet.write('N1', 'Phí VAT')
worksheet.write('O1', 'Tổng tiền')
worksheet.write('P1', 'Số hóa đơn')
worksheet.write('Q1', 'Seri hóa đơn')
worksheet.write('R1', 'Ngày lập HĐ')
worksheet.write('S1', 'Ngày đầu kỳ')
worksheet.write('T1', 'ngày cuối kỳ')
worksheet.write('U1', 'Đã thanh toán')
worksheet.write('V1', 'Người thu tiền')
worksheet.write('W1', 'Ghi chú')

ID = driver.find_elements(By.CSS_SELECTOR,'div.col_full:nth-child(1)>input')
ID[0].send_keys("AT_Test")
time.sleep(1)
PASS = driver.find_elements(By.CSS_SELECTOR,'div.col_full:nth-child(2)>input')
PASS[0].send_keys("1234567890Aa")
BTN = driver.find_elements(By.ID,'ctl00_mainContent_login1_LoginCtrl_Login')
BTN[0].click()
driver.get("https://app.citywork.vn/gcrm.aspx")

time.sleep(1)

driver.find_element(By.ID,'ext-gen1171').click()
time.sleep(5)
driver.find_elements(By.CSS_SELECTOR,'li.x-boundlist-item')[4].click()
time.sleep(1)
driver.find_element(By.ID,'ext-gen1215').click()
time.sleep(2)
thang = driver.find_element(By.NAME,'thang')
count = 1
for i in cacThang:
    thang.clear()
    for c in i:
        thang.send_keys(c)
    time.sleep(5)
    maxPage = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
    while(maxPage>4000):        
        thang.clear()
        thang.send_keys("08/2023")
        thang.clear()
        for c in i:
            thang.send_keys(c)
        time.sleep(5)
        maxPage = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
    print(i+ " có "+str(maxPage)+" trang")
    page = 1
    while(page != maxPage):
        table = driver.find_elements(By.CSS_SELECTOR,"table.x-grid-table")
        stt = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(1)")
        soHopDong = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(6)")
        maHopDong = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(7)")
        tuyenDoc = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(8)")
        tenKhachHang = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(9)")
        diaChi = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(10)")
        chiSoCu = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(13)")
        chiSoMoi = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(14)")
        tieuThu = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(16)")
        maDTGia = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(17)")
        thanhTien = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(18)")
        phiDTDN = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(19)")
        phiBVMT = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(20)")
        phiVAT = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(23)")
        tongTien = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(25)")
        soHoaDon = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(27)")
        seriHoaDon = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(28)")
        ngayLapHoaDon = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(29)")
        ngayHD = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(30)")
        ngayDauKy = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(31)")
        ngayCuoiKy = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(32)")
        daThanhToan = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(33)")
        nguoiThuTien = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(34)")
        ghiChu = table[3].find_elements(By.CSS_SELECTOR,"tr.x-grid-row > td:nth-child(35)")
        for index in range(len(stt)):
            worksheet.write(count, 0, stt[index].text)
            worksheet.write(count, 1, soHopDong[index].text)
            worksheet.write(count, 2, maHopDong[index].text)
            worksheet.write(count, 3, tuyenDoc[index].text)
            worksheet.write(count, 4, tenKhachHang[index].text)
            worksheet.write(count, 5, diaChi[index].text)
            worksheet.write(count, 6, chiSoCu[index].text)
            worksheet.write(count, 7, chiSoMoi[index].text)
            worksheet.write(count, 8, tieuThu[index].text)
            worksheet.write(count, 9, maDTGia[index].text)
            worksheet.write(count, 10, thanhTien[index].text)
            worksheet.write(count, 11, phiDTDN[index].text)
            worksheet.write(count, 12, phiBVMT[index].text)
            worksheet.write(count, 13, phiVAT[index].text)
            worksheet.write(count, 14, tongTien[index].text)
            worksheet.write(count, 15, soHoaDon[index].text)
            worksheet.write(count, 16, seriHoaDon[index].text)
            worksheet.write(count, 17, ngayLapHoaDon[index].text)
            worksheet.write(count, 18, ngayHD[index].text)
            worksheet.write(count, 19, ngayDauKy[index].text)
            worksheet.write(count, 20, ngayCuoiKy[index].text)
            worksheet.write(count, 21, daThanhToan[index].text)
            worksheet.write(count, 22, nguoiThuTien[index].text)
            worksheet.write(count, 23, ghiChu[index].text)
            print(stt[index].text+" "+ tenKhachHang[index].text)
            print("----------------------")
            count +=1
        page = int(driver.find_element(By.NAME,'inputItem').get_attribute('value'))
        driver.find_element(By.CSS_SELECTOR,'span.x-tbar-page-next').click()
        time.sleep(2)
        print('---------------Trang '+str(page)+' hoàn thành---------------')
    print("---------------Tháng "+i + " hoàn thành---------------")
workbook.close()
driver.quit()
end = time.time()
print(str(end - start)+"s")