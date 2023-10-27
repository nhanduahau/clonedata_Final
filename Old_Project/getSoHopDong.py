from asyncio import exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
import mouse
from selenium.webdriver.common.action_chains import ActionChains
import xlsxwriter
import time

STARPAGE = 1
ENDPAGE = 5 #max 174



workbook = xlsxwriter.Workbook("hopDong"+str(STARPAGE)+"-"+str(ENDPAGE)+".xlsx")
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Mã khách hàng')
worksheet.write('B1', 'Tên khách hàng')
worksheet.write('C1', 'Địa chỉ')
worksheet.write('D1', 'Tên thường gọi')
worksheet.write('E1', 'Số hộ')
worksheet.write('F1', 'Số khẩu')
worksheet.write('G1', 'Email')
worksheet.write('H1', 'Ngày cấp CMND')
worksheet.write('I1', 'Nơi cấp CMND')
worksheet.write('J1', 'Mã số thuế')
worksheet.write('K1', 'Người đại điện')
worksheet.write('L1', 'Đối tượng')
worksheet.write('M1', 'Ghi chú')
worksheet.write('N1', 'Số hợp đồng')
worksheet.write('O1', 'ĐT giá')
worksheet.write('P1', 'Mục đích SD')
worksheet.write('Q1', 'Hình thức thanh toán')
worksheet.write('R1', 'Mã vạch')
worksheet.write('S1', 'Ngày ký')
worksheet.write('T1', 'Ngày lắp đặt')
worksheet.write('U1', 'Người lắp đặt')
worksheet.write('V1', 'ngàyNT')
worksheet.write('W1', 'Tiền lắp đặt')
worksheet.write('X1', 'Người nộp')
worksheet.write('Y1', 'Tiền đặt cọc')
worksheet.write('Z1', 'Giảm trừ theo')
worksheet.write('AA1', 'Số tiền giảm trừ')
worksheet.write('AB1', 'Ngày đặt cọc')
worksheet.write('AC1', 'Chứng từ đặt cọc')
worksheet.write('AD1', 'Cam kết sử dụng nước')
worksheet.write('AE1', 'Khối lượng cam kết')
worksheet.write('AF1', 'Ghi chú hợp đồng')

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


driver.find_element(By.ID,'ext-gen1202').click()
time.sleep(2)
def save():
    worksheet.write(count, 0, maKhachHang)
    worksheet.write(count, 1, tenKhachHang)
    worksheet.write(count, 2, diaChi1)
    worksheet.write(count, 3, tenThuongGoi)
    worksheet.write(count, 4, soHoDungChung)
    worksheet.write(count, 5, soNhanKhau)
    worksheet.write(count, 6, email)
    worksheet.write(count, 7, ngayCapCMND)
    worksheet.write(count, 8, noiCapCMND)
    worksheet.write(count, 9, maSoThue)
    worksheet.write(count, 10, nguoiDaiDien)
    worksheet.write(count, 11, doiTuong)
    worksheet.write(count, 12, ghiChu)
    worksheet.write(count, 13, soHopDong)
    worksheet.write(count, 14, maDoiTuongGia)
    worksheet.write(count, 15, mucDich)
    worksheet.write(count, 16, maPhuongThucThanhToan)
    worksheet.write(count, 17, maVach)
    worksheet.write(count, 18, ngayKy)
    worksheet.write(count, 19, ngayLapDat)
    worksheet.write(count, 20, nguoiLapDat)
    worksheet.write(count, 21, ngayBanGiao)
    worksheet.write(count, 22, soTien)
    worksheet.write(count, 23, nguoiNop)
    worksheet.write(count, 24, tienDatCoc)
    worksheet.write(count, 25, loaiGiamTru)
    worksheet.write(count, 26, soTienGiamTru)
    worksheet.write(count, 27, ngayDatCoc)
    worksheet.write(count, 28, chungTuDatCoc)
    worksheet.write(count, 29, camKetSuDungNuoc)
    worksheet.write(count, 30, khoiLuongCamKet)
    worksheet.write(count, 31, ghiChuHD)
def isCheck(table,ch):
    cam_ket = table.find_elements(By.CSS_SELECTOR,"input.x-form-checkbox")[9]
    dong_ho_phu = table.find_elements(By.CSS_SELECTOR,"input.x-form-checkbox")[10]
    tick = table.find_elements(By.CSS_SELECTOR,"table.x-form-cb-checked")
    for t in tick:
        for k in t.find_elements(By.CSS_SELECTOR,"input.x-form-checkbox"):
            if(k==dong_ho_phu and ch=="DHP"):
                return "có"
            if(k==cam_ket and ch=="CK"):
                return "có"
    return "không"
count = 1
driver.find_element(By.NAME,'inputItem').clear()
driver.find_element(By.NAME,'inputItem').send_keys(STARPAGE)
driver.find_element(By.NAME,'inputItem').send_keys(Keys.ENTER)
page =int(driver.find_element(By.NAME,'inputItem').get_attribute('value'))
time.sleep(1)
while(page!=ENDPAGE+1):
    table = driver.find_elements(By.CSS_SELECTOR,'table.x-grid-table')[3]
    row = table.find_elements(By.CSS_SELECTOR,'tr.x-grid-row')
    stt = 1
    for i in range(len(row)):
        if(count > 200):
            break
        row[i].click()
        time.sleep(0.5)
        driver.find_elements(By.CSS_SELECTOR,'span.x-btn-inner')[13].click()
        time.sleep(0.5)
        table = driver.find_element(By.CSS_SELECTOR,"div.x-window-default")
        data = table.find_elements(By.CSS_SELECTOR,'table.x-grid-table > tbody > tr.x-grid-row')
        maKhachHang = table.find_element(By.NAME,"maKhachHang").get_attribute("value")
        tenKhachHang = table.find_element(By.NAME,"tenKhachHang").get_attribute("value")
        diaChi1 = table.find_elements(By.NAME,"diaChi")[0].get_attribute("value")
        tenThuongGoi = table.find_element(By.NAME,"tenThuongGoi").get_attribute("value")
        soHoDungChung = table.find_element(By.NAME,"soHoDungChung").get_attribute("value")
        soNhanKhau = table.find_element(By.NAME,"soNhanKhau").get_attribute("value")
        email = table.find_element(By.NAME,"email").get_attribute("value")
        ngayCapCMND = table.find_element(By.NAME,"ngayCapCMND").get_attribute("value")
        noiCapCMND = table.find_element(By.NAME,"noiCapCMND").get_attribute("value")
        maSoThue = table.find_element(By.NAME,"maSoThue").get_attribute("value")
        nguoiDaiDien = table.find_element(By.NAME,"nguoiDaiDien").get_attribute("value")
        doiTuong = table.find_element(By.NAME,"doiTuong").get_attribute("value")
        ghiChu = table.find_element(By.NAME,"ghiChu").get_attribute("value")
        soHopDong = table.find_element(By.NAME,"soHopDong").get_attribute("value")
        maDoiTuongGia = table.find_element(By.NAME,"maDoiTuongGia").get_attribute("value")
        mucDich = table.find_element(By.NAME,"mucDich").get_attribute("value")
        maPhuongThucThanhToan = table.find_element(By.NAME,"maPhuongThucThanhToan").get_attribute("value")
        maVach = table.find_element(By.NAME,"maVach").get_attribute("value")
        ngayKy = table.find_element(By.NAME,"ngayKy").get_attribute("value")
        ngayLapDat = table.find_element(By.NAME,"ngayLapDat").get_attribute("value")
        nguoiLapDat = table.find_element(By.NAME,"nguoiLapDat").get_attribute("value")
        ngayBanGiao = table.find_element(By.NAME,"ngayBanGiao").get_attribute("value")
        soTien = table.find_element(By.NAME,"soTien").get_attribute("value")
        nguoiNop = table.find_element(By.NAME,"nguoiNop").get_attribute("value")
        tienDatCoc = table.find_element(By.NAME,"tienDatCoc").get_attribute("value")
        loaiGiamTru = table.find_element(By.NAME,"loaiGiamTru").get_attribute("value")
        soTienGiamTru = table.find_element(By.NAME,"soTienGiamTru").get_attribute("value")
        ngayDatCoc = table.find_element(By.NAME,"ngayDatCoc").get_attribute("value")
        chungTuDatCoc = table.find_element(By.NAME,"chungTuDatCoc").get_attribute("value")
        camKetSuDungNuoc = isCheck(table,"CK")
        khoiLuongCamKet = table.find_element(By.NAME,"khoiLuongCamKet").get_attribute("value")
        xDaiDien = table.find_element(By.NAME,"xDaiDien").get_attribute("value")
        yDaiDien = table.find_element(By.NAME,"yDaiDien").get_attribute("value")
        ghiChuHD  = table.find_element(By.NAME,"ghiChuHD").get_attribute("value")
        save()
        count +=1
        print(str(stt)+ " -- " +tenKhachHang)
        table.find_element(By.CSS_SELECTOR,"span.iconClose").click()
        stt +=1
        
    driver.find_element(By.CSS_SELECTOR,'span.x-tbar-page-next').click()
    time.sleep(0.5)
    print('---------------Trang '+str(page)+' hoàn thành---------------')
    page +=1
    
workbook.close()
driver.quit()