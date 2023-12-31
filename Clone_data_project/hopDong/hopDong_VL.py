from asyncio import exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
import mouse
from selenium.webdriver.common.action_chains import ActionChains
import xlsxwriter
import time

#------------------------------------VL--------------------------------------------
nameOfOutputFileExcel ="Clone_data_project/hopDong/results/hopDong_VL.xlsx"
nameOfOutputFileText = "Clone_data_project/hopDong/results/soHopDong_VL.txt"
#--------------------------------------------------------------------------------

workbook = xlsxwriter.Workbook(nameOfOutputFileExcel)
worksheet = workbook.add_worksheet()
worksheet.write('A1 ', 'Loại khách hàng')     
worksheet.write('B1 ', 'Mã khách hàng')     
worksheet.write('C1 ', 'Tên khách hàng')     
worksheet.write('D1 ', 'Địa chỉ')     
worksheet.write('E1 ', 'Tên thường gọi')     
worksheet.write('F1 ', 'Số hộ')     
worksheet.write('G1 ', 'Số khẩu')     
worksheet.write('H1 ', 'Email')     
worksheet.write('I1 ', 'Nhà mạng')     
worksheet.write('J1 ', 'Số điện thoại')     
worksheet.write('K1 ', 'Số CMND')     
worksheet.write('L1 ', 'Ngày cấp CMND')     
worksheet.write('M1 ', 'Nơi cấp CMND')     
worksheet.write('N1 ', 'Mã số thuế')     
worksheet.write('O1 ', 'Tên ngân hàng')     
worksheet.write('P1 ', 'Số GCN')     
worksheet.write('Q1 ', 'Tên TKNH')     
worksheet.write('R1 ', 'Số TKNH')     
worksheet.write('S1 ', 'Nguồn nước')     
worksheet.write('T1 ', 'Người đại điện')     
worksheet.write('U1 ', 'Ghi chú')     
worksheet.write('V1 ', 'Đối tượng')     
worksheet.write('W1 ', 'Mã đăng ký')    
worksheet.write('X1 ', 'Số hợp đồng')    
worksheet.write('Y1 ', 'ĐT giá')    
worksheet.write('Z1 ', 'Mục đích SD')    
worksheet.write('AA1', 'Khu vực TT')    
worksheet.write('AB1', 'Hình thức TT')    
worksheet.write('AC1', 'Mã vạch')    
worksheet.write('AD1', 'Ngày ký HĐ')    
worksheet.write('AE1', 'Ngày lắp đặt')    
worksheet.write('AF1', 'Người lắp đặt')    
worksheet.write('AG1', 'Ngày NT')    
worksheet.write('AH1', 'Tiền lắp đặt')    
worksheet.write('AI1', 'Người nộp')    
worksheet.write('AJ1', 'Tiền đặt cọc')    
worksheet.write('AK1', 'Giảm trừ theo')    
worksheet.write('AL1', 'Số tiền giảm trừ')    
worksheet.write('AM1', 'Ngày đặt cọc')    
worksheet.write('AN1', 'Chứng từ đặt cọc')    
worksheet.write('AO1', 'Cam kết sử dụng nước')    
worksheet.write('AP1', 'Khối lượng cam kết')    
worksheet.write('AQ1', 'Ghi chú hợp đồng')    
worksheet.write('AR1', 'Tỉnh')    
worksheet.write('AS1', 'Huyện')    
worksheet.write('AT1', 'Xã')    
worksheet.write('AU1', 'Vùng')    
worksheet.write('AV1', 'Khu vực')    
worksheet.write('AW1', 'Nhân viên')    
worksheet.write('AX1', 'Tuyến đọc')    
worksheet.write('AY1', 'Phạm vi')    
worksheet.write('AZ1', 'Mã đồng hồ')    
worksheet.write('BA1', 'Đồng hồ block')    
worksheet.write('BB1', 'Là đồng hồ phụ')    
worksheet.write('BC1', 'Số thứ tự')    
worksheet.write('BD1', 'Seri')    
worksheet.write('BE1', 'Chỉ số đầu')    
worksheet.write('BF1', 'Chỉ số cuối')    
worksheet.write('BG1', 'Seri chỉ')    
worksheet.write('BH1', 'Ngày lắp đặt đồng hồ')    
worksheet.write('BI1', 'Ngày sử dụng')    
worksheet.write('BJ1', 'Địa chỉ')    
worksheet.write('BK1', 'Trạng thái')    
worksheet.write('BL1', 'Lý do hủy')    
worksheet.write('BM1','Ngày BĐ ngừng')    
worksheet.write('BN1','Ngày kết thúc')    
worksheet.write('BO1', 'Nước SX')    
worksheet.write('BP1', 'Hãng SX')    
worksheet.write('BQ1', 'Kiểu đồng hồ')    
worksheet.write('BR1', 'Đường kính')    
worksheet.write('BS1', 'Hộp bảo vệ')    
worksheet.write('BT1', 'Vị trí lắp đặt')    
worksheet.write('BU1', 'Ngày kiểm định')    
worksheet.write('BV1', 'Hiệu lực KĐ')    
worksheet.write('BW1', 'Lý do kiểm định')    
worksheet.write('BX1', 'Van một chiều')    
worksheet.write('BY1', 'Số Tem')    
worksheet.write('BZ1', 'Số phiếu thay')    
worksheet.write('CA1', 'Hình thức XL')    
worksheet.write('CB1', 'Lý do thay')    
worksheet.write('CC1', 'Mã ĐH thay')    
worksheet.write('CD1', 'Người thay')    
worksheet.write('CE1', 'Kinh độ')    
worksheet.write('CF1', 'Vĩ độ')    
worksheet.write('CG1', 'Lọa KM')    
worksheet.write('CH1', 'Khuyến mãi')    
worksheet.write('CI1', 'Trạng thái ĐH lắp')    
worksheet.write('CJ1', 'Ống dẫn')    
worksheet.write('CK1', 'Đai khởi thủy') 

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


driver.find_element(By.XPATH,'/html/body/div[2]/div[4]/div/table/tbody/tr[10]/td/div').click()
time.sleep(2)
STARPAGE = 1
ENDPAGE = int(driver.find_elements(By.CSS_SELECTOR,'div.x-toolbar-text-default')[1].text.replace("của ",""))
def save():
    maTinh  = table.find_element(By.NAME,"maTinh").get_attribute("value")
    maHuyen = table.find_element(By.NAME,"maHuyen").get_attribute("value")
    maXa = table.find_element(By.NAME,"maXa").get_attribute("value")
    maDMVung = table.find_element(By.NAME,"maDMVung").get_attribute("value")
    maDMKhuVuc = table.find_element(By.NAME,"maDMKhuVuc").get_attribute("value")
    maNhanVien = table.find_element(By.NAME,"maNhanVien").get_attribute("value")
    maTuyenDoc = table.find_element(By.NAME,"maTuyenDoc").get_attribute("value")
    maPhamVi = table.find_element(By.NAME,"maPhamVi").get_attribute("value")
    maDongHo = table.find_element(By.NAME,"maDongHo").get_attribute("value")
    maDongHoCha = table.find_element(By.NAME,"maDongHoCha").get_attribute("value")
    laDongHoPhu = isCheck(table,"DHP")
    soThuTu = table.find_element(By.NAME,"soThuTu").get_attribute("value")
    seri = table.find_element(By.NAME,"seri").get_attribute("value")
    chiSoDau = table.find_element(By.NAME,"chiSoDau").get_attribute("value")
    chiSoCuoi = table.find_element(By.NAME,"chiSoCuoi").get_attribute("value")
    seriChi = table.find_element(By.NAME,"seriChi").get_attribute("value")
    ngayLapDatDongHo = table.find_element(By.NAME,"ngayLapDatDongHo").get_attribute("value")
    ngayBatDauSuDung = table.find_element(By.NAME,"ngayBatDauSuDung").get_attribute("value")
    diaChi2 = table.find_elements(By.NAME,"diaChi")[1].get_attribute("value")
    maTinhTrangDongHo = table.find_element(By.NAME,"maTinhTrangDongHo").get_attribute("value")
    lyDoHuy = table.find_element(By.NAME,"lyDoHuy").get_attribute("value")
    ngayBatDauNgung = table.find_element(By.NAME,"ngayBatDauNgung").get_attribute("value")
    ngayKetThucSuDung = table.find_element(By.NAME,"ngayKetThucSuDung").get_attribute("value")
    maDMNuocSanXuat = table.find_element(By.NAME,"maDMNuocSanXuat").get_attribute("value")
    maDMHangSanXuat = table.find_element(By.NAME,"maDMHangSanXuat").get_attribute("value")
    maDMKieuDongHo = table.find_element(By.NAME,"maDMKieuDongHo").get_attribute("value")
    duongKinh = table.find_element(By.NAME,"duongKinh").get_attribute("value")
    maDMHopBaoVeDongHo = table.find_element(By.NAME,"maDMHopBaoVeDongHo").get_attribute("value")
    viTriLapDat = table.find_element(By.NAME,"viTriLapDat").get_attribute("value")
    ngayKiemDinh = table.find_element(By.NAME,"ngayKiemDinh").get_attribute("value")
    ngayHieuLucKiemDinh = table.find_element(By.NAME,"ngayHieuLucKiemDinh").get_attribute("value")
    lyDoKiemDinh = table.find_element(By.NAME,"lyDoKiemDinh").get_attribute("value")
    vanMotChieu = table.find_element(By.NAME,"vanMotChieu").get_attribute("value")
    tem = table.find_element(By.NAME,"tem").get_attribute("value")
    soPhieuThay = table.find_element(By.NAME,"soPhieuThay").get_attribute("value")
    hinhThucXuLy = table.find_element(By.NAME,"hinhThucXuLy").get_attribute("value")
    lyDoThay = table.find_element(By.NAME,"lyDoThay").get_attribute("value")
    maDongHoThay = table.find_element(By.NAME,"maDongHoThay").get_attribute("value")
    nguoiThay = table.find_element(By.NAME,"nguoiThay").get_attribute("value")
    xDaiDien = table.find_element(By.NAME,"xDaiDien").get_attribute("value")
    yDaiDien = table.find_element(By.NAME,"yDaiDien").get_attribute("value")
    loaiKhuyenMai = table.find_element(By.NAME,"loaiKhuyenMai").get_attribute("value")
    soNuocKhuyenMaiConLai = table.find_element(By.NAME,"soNuocKhuyenMaiConLai").get_attribute("value")
    trangThaiDongHo = table.find_element(By.NAME,"trangThaiDongHo").get_attribute("value")
    ongDan = table.find_element(By.NAME,"ongDan").get_attribute("value")
    daiKhoiThuy = table.find_element(By.NAME,"daiKhoiThuy").get_attribute("value")
    worksheet.write(count, 0 , maLoaiKhachHang)
    worksheet.write(count, 1 , maKhachHang)
    worksheet.write(count, 2 , tenKhachHang)
    worksheet.write(count, 3 , diaChi1)
    worksheet.write(count, 4 , tenThuongGoi)
    worksheet.write(count, 5 , soHoDungChung)
    worksheet.write(count, 6 , soNhanKhau)
    worksheet.write(count, 7 , email)
    worksheet.write(count, 8 ,tenNhaMang)
    worksheet.write(count, 9 ,dienThoai)
    worksheet.write(count, 10,soCMND)
    worksheet.write(count, 11, ngayCapCMND)
    worksheet.write(count, 12, noiCapCMND)
    worksheet.write(count, 13, maSoThue)
    worksheet.write(count, 14,tenNganHang)
    worksheet.write(count, 15,soGCN)
    worksheet.write(count, 16,tenTaiKhoanNH)
    worksheet.write(count, 17,taiKhoanNganHang)
    worksheet.write(count, 18,nguonNuocKhac)
    worksheet.write(count, 19, nguoiDaiDien)
    worksheet.write(count, 20, ghiChu)
    worksheet.write(count, 21, doiTuong)
    worksheet.write(count, 22, maDangKyLapDat)
    worksheet.write(count, 23, soHopDong)
    worksheet.write(count, 24, maDoiTuongGia)
    worksheet.write(count, 25, mucDich)
    worksheet.write(count, 26,maKhuVucThanhToan)
    worksheet.write(count, 27, maPhuongThucThanhToan)
    worksheet.write(count, 28, maVach)
    worksheet.write(count, 29, ngayKy)
    worksheet.write(count, 30, ngayLapDat)
    worksheet.write(count, 31, nguoiLapDat)
    worksheet.write(count, 32, ngayBanGiao)
    worksheet.write(count, 33, soTien)
    worksheet.write(count, 34, nguoiNop)
    worksheet.write(count, 35, tienDatCoc)
    worksheet.write(count, 36, loaiGiamTru)
    worksheet.write(count, 37, soTienGiamTru)
    worksheet.write(count, 38, ngayDatCoc)
    worksheet.write(count, 39, chungTuDatCoc)
    worksheet.write(count, 40, camKetSuDungNuoc)
    worksheet.write(count, 41, khoiLuongCamKet)
    worksheet.write(count, 42, ghiChuHD)
    worksheet.write(count, 43, maTinh)
    worksheet.write(count, 44, maHuyen)
    worksheet.write(count, 45, maXa)
    worksheet.write(count, 46, maDMVung)
    worksheet.write(count, 47, maDMKhuVuc)
    worksheet.write(count, 48, maNhanVien)
    worksheet.write(count, 49, maTuyenDoc)
    worksheet.write(count, 50, maPhamVi)
    worksheet.write(count, 51, maDongHo)
    worksheet.write(count, 52, maDongHoCha)
    worksheet.write(count, 53, laDongHoPhu)
    worksheet.write(count, 54, soThuTu)
    worksheet.write(count, 55,seri)
    worksheet.write(count, 56, chiSoDau)
    worksheet.write(count, 57, chiSoCuoi)
    worksheet.write(count, 58, seriChi)
    worksheet.write(count, 59, ngayLapDatDongHo)
    worksheet.write(count, 60, ngayBatDauSuDung)
    worksheet.write(count, 61, diaChi2)
    worksheet.write(count, 62, maTinhTrangDongHo)
    worksheet.write(count, 63, lyDoHuy)
    worksheet.write(count, 64, ngayBatDauNgung)
    worksheet.write(count, 65, ngayKetThucSuDung)
    worksheet.write(count, 66, maDMNuocSanXuat)
    worksheet.write(count, 67, maDMHangSanXuat)
    worksheet.write(count, 68, maDMKieuDongHo)
    worksheet.write(count, 69, duongKinh)
    worksheet.write(count, 70, maDMHopBaoVeDongHo)
    worksheet.write(count, 71, viTriLapDat)
    worksheet.write(count, 72, ngayKiemDinh)
    worksheet.write(count, 73, ngayHieuLucKiemDinh)
    worksheet.write(count, 74, lyDoKiemDinh)
    worksheet.write(count, 75, vanMotChieu)
    worksheet.write(count, 76, tem)
    worksheet.write(count, 77, soPhieuThay)
    worksheet.write(count, 78, hinhThucXuLy)
    worksheet.write(count, 79, lyDoThay)
    worksheet.write(count, 80, maDongHoThay)
    worksheet.write(count, 81, nguoiThay)
    worksheet.write(count, 82, xDaiDien)
    worksheet.write(count, 83, yDaiDien)
    worksheet.write(count, 84, loaiKhuyenMai)
    worksheet.write(count, 85, soNuocKhuyenMaiConLai)
    worksheet.write(count, 86, trangThaiDongHo)
    worksheet.write(count, 87, ongDan)
    worksheet.write(count, 88, daiKhoiThuy)
    
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
stt = 1

while(page!=ENDPAGE+1):
    table = driver.find_elements(By.CSS_SELECTOR,'table.x-grid-table')[3]
    row = table.find_elements(By.CSS_SELECTOR,'tr.x-grid-row')
    for i in range(len(row)):
        row[i].click()
        time.sleep(1)
        driver.find_elements(By.CSS_SELECTOR,'span.x-btn-inner')[13].click()
        time.sleep(1)
        table = driver.find_element(By.CSS_SELECTOR,"div.x-window-default")
        data = table.find_elements(By.CSS_SELECTOR,'table.x-grid-table > tbody > tr.x-grid-row')
        maLoaiKhachHang = table.find_element(By.NAME,"maLoaiKhachHang").get_attribute("value")
        maKhachHang = table.find_element(By.NAME,"maKhachHang").get_attribute("value")
        tenKhachHang = table.find_element(By.NAME,"tenKhachHang").get_attribute("value")
        diaChi1 = table.find_elements(By.NAME,"diaChi")[0].get_attribute("value")
        tenThuongGoi = table.find_element(By.NAME,"tenThuongGoi").get_attribute("value")
        soHoDungChung = table.find_element(By.NAME,"soHoDungChung").get_attribute("value")
        soNhanKhau = table.find_element(By.NAME,"soNhanKhau").get_attribute("value")
        email = table.find_element(By.NAME,"email").get_attribute("value")
        tenNhaMang = table.find_element(By.NAME,"tenNhaMang").get_attribute("value")
        dienThoai = table.find_element(By.NAME,"dienThoai").get_attribute("value")
        soCMND = table.find_element(By.NAME,"soCMND").get_attribute("value")
        ngayCapCMND = table.find_element(By.NAME,"ngayCapCMND").get_attribute("value")
        noiCapCMND = table.find_element(By.NAME,"noiCapCMND").get_attribute("value")
        maSoThue = table.find_element(By.NAME,"maSoThue").get_attribute("value")
        tenNganHang = table.find_element(By.NAME,"tenNganHang").get_attribute("value")
        soGCN = table.find_element(By.NAME,"soGCN").get_attribute("value")
        tenTaiKhoanNH = table.find_element(By.NAME,"tenTaiKhoanNH").get_attribute("value")
        taiKhoanNganHang = table.find_element(By.NAME,"taiKhoanNganHang").get_attribute("value")
        nguonNuocKhac = table.find_element(By.NAME,"nguonNuocKhac").get_attribute("value")
        nguoiDaiDien = table.find_element(By.NAME,"nguoiDaiDien").get_attribute("value")
        doiTuong = table.find_element(By.NAME,"doiTuong").get_attribute("value")
        ghiChu = table.find_element(By.NAME,"ghiChu").get_attribute("value")
        maDangKyLapDat = table.find_element(By.NAME,"maDangKyLapDat").get_attribute("value")
        soHopDong = table.find_element(By.NAME,"soHopDong").get_attribute("value")
        maDoiTuongGia = table.find_element(By.NAME,"maDoiTuongGia").get_attribute("value")
        mucDich = table.find_element(By.NAME,"mucDich").get_attribute("value")
        maKhuVucThanhToan = table.find_element(By.NAME,"maKhuVucThanhToan").get_attribute("value")
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
        maLoaiKhachHang = table.find_element(By.NAME,"maLoaiKhachHang").get_attribute("value")
        for j in range(len(data)):
            try:
                loc = data[j].location
                data[j].click()
                time.sleep(1)
            except Exception as e:
                loc["y"] += 27
                time.sleep(1)
                mouse.move(int(loc["x"]+5), int(loc["y"]+137))
                mouse.click(button='left')
                time.sleep(1)
            save()
            count +=1
        f = open(nameOfOutputFileText, "a")
        f.write(soHopDong+"\n")
        f.close()
        print(str(stt)+ " -- " +tenKhachHang)
        table.find_element(By.CSS_SELECTOR,"span.iconClose").click()
        stt +=1
    driver.find_element(By.CSS_SELECTOR,'span.x-tbar-page-next').click()
    time.sleep(1)
    print('---------------Trang '+str(page)+' hoàn thành---------------')
    page +=1
workbook.close()
driver.quit()