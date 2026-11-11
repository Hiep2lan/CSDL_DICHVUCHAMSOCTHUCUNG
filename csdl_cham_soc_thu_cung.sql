-- =================================================================================
-- HỆ THỐNG CƠ SỞ DỮ LIỆU QUẢN LÝ DỊCH VỤ CHĂM SÓC THÚ CƯNG (ĐÃ CHUẨN HÓA 3NF)
-- Bao gồm: DDL khởi tạo bảng, Ràng buộc toàn vẹn, Dữ liệu mẫu, và SQL cho Mục 5, 6, 7
-- =================================================================================

PRAGMA foreign_keys = ON;

DROP TABLE IF EXISTS THANH_TOAN;
DROP TABLE IF EXISTS LICH_HEN;
DROP TABLE IF EXISTS THU_CUNG;
DROP TABLE IF EXISTS DICH_VU;
DROP TABLE IF EXISTS NHAN_VIEN;
DROP TABLE IF EXISTS KHACH_HANG;

-- 1. BẢNG KHACH_HANG (Đạt 3NF)
CREATE TABLE KHACH_HANG (
    MaKH VARCHAR(10) PRIMARY KEY,
    HoTen NVARCHAR(100) NOT NULL,
    SDT VARCHAR(15) NOT NULL UNIQUE,
    Email VARCHAR(100) UNIQUE,
    DiaChi NVARCHAR(255),
    CONSTRAINT CHK_KH_SDT CHECK (LENGTH(SDT) BETWEEN 10 AND 11 AND SDT GLOB '0[0-9]*')
);

-- 2. BẢNG THU_CUNG (Đạt 3NF)
CREATE TABLE THU_CUNG (
    MaThuCung VARCHAR(10) PRIMARY KEY,
    MaKH VARCHAR(10) NOT NULL,
    TenThuCung NVARCHAR(100) NOT NULL,
    Loai NVARCHAR(50) NOT NULL,
    Giong NVARCHAR(100),
    GioiTinh NVARCHAR(20) NOT NULL DEFAULT 'Chưa xác định',
    Tuoi INT NOT NULL,
    CanNang DECIMAL(5,2) NOT NULL,
    GhiChu NVARCHAR(500),
    CONSTRAINT FK_TC_KH FOREIGN KEY (MaKH) REFERENCES KHACH_HANG(MaKH) ON UPDATE CASCADE ON DELETE RESTRICT,
    CONSTRAINT CHK_TC_GioiTinh CHECK (GioiTinh IN ('Đực', 'Cái', 'Chưa xác định')),
    CONSTRAINT CHK_TC_Tuoi CHECK (Tuoi > 0 AND Tuoi <= 360),
    CONSTRAINT CHK_TC_CanNang CHECK (CanNang > 0 AND CanNang <= 150)
);

-- 3. BẢNG NHAN_VIEN (Đạt 3NF)
CREATE TABLE NHAN_VIEN (
    MaNV VARCHAR(10) PRIMARY KEY,
    HoTen NVARCHAR(100) NOT NULL,
    NgaySinh DATE NOT NULL,
    GioiTinh NVARCHAR(10) NOT NULL,
    SDT VARCHAR(15) NOT NULL UNIQUE,
    ChucVu NVARCHAR(50) NOT NULL,
    CaLamViec NVARCHAR(50) NOT NULL,
    Luong DECIMAL(12,2) NOT NULL,
    TrangThai NVARCHAR(50) NOT NULL DEFAULT 'Đang làm việc',
    CONSTRAINT CHK_NV_GioiTinh CHECK (GioiTinh IN ('Nam', 'Nữ')),
    CONSTRAINT CHK_NV_ChucVu CHECK (ChucVu IN ('Lễ tân', 'Nhân viên chăm sóc', 'Quản lý')),
    CONSTRAINT CHK_NV_CaLamViec CHECK (CaLamViec IN ('Ca sáng (08:00-12:00)', 'Ca chiều (13:00-17:00)', 'Ca tối (17:00-21:00)', 'Hành chính (08:00-17:00)')),
    CONSTRAINT CHK_NV_Luong CHECK (Luong >= 0),
    CONSTRAINT CHK_NV_TrangThai CHECK (TrangThai IN ('Đang làm việc', 'Nghỉ phép', 'Đã nghỉ việc'))
);

-- 4. BẢNG DICH_VU (Đạt 3NF)
CREATE TABLE DICH_VU (
    MaDV VARCHAR(10) PRIMARY KEY,
    TenDV NVARCHAR(150) NOT NULL UNIQUE,
    DonGia DECIMAL(12,2) NOT NULL,
    ThoiLuongDuKien INT NOT NULL, -- Tính bằng phút
    MoTaDichVu NVARCHAR(500),
    CONSTRAINT CHK_DV_DonGia CHECK (DonGia >= 0),
    CONSTRAINT CHK_DV_ThoiLuong CHECK (ThoiLuongDuKien > 0)
);

-- 5. BẢNG LICH_HEN (Đã chuẩn hóa 3NF: Loại bỏ MaKH dư thừa bắc cầu qua MaThuCung)
CREATE TABLE LICH_HEN (
    MaLich VARCHAR(10) PRIMARY KEY,
    MaThuCung VARCHAR(10) NOT NULL,
    MaDV VARCHAR(10) NOT NULL,
    MaNV VARCHAR(10) NOT NULL,
    NgayGioHen DATETIME NOT NULL,
    SoLuong INT NOT NULL DEFAULT 1,
    TrangThai NVARCHAR(50) NOT NULL DEFAULT 'Mới tạo',
    GhiChu NVARCHAR(500),
    CONSTRAINT FK_LH_TC FOREIGN KEY (MaThuCung) REFERENCES THU_CUNG(MaThuCung) ON UPDATE CASCADE ON DELETE RESTRICT,
    CONSTRAINT FK_LH_DV FOREIGN KEY (MaDV) REFERENCES DICH_VU(MaDV) ON UPDATE CASCADE ON DELETE RESTRICT,
    CONSTRAINT FK_LH_NV FOREIGN KEY (MaNV) REFERENCES NHAN_VIEN(MaNV) ON UPDATE CASCADE ON DELETE RESTRICT,
    CONSTRAINT CHK_LH_SoLuong CHECK (SoLuong > 0),
    CONSTRAINT CHK_LH_TrangThai CHECK (TrangThai IN ('Mới tạo', 'Đã xác nhận', 'Đang thực hiện', 'Hoàn thành', 'Đã hủy', 'Không đến'))
);

-- 6. BẢNG THANH_TOAN (Đã chuẩn hóa 3NF: Loại bỏ ThanhTien dẫn xuất, kết nối 1-1 với LICH_HEN)
CREATE TABLE THANH_TOAN (
    MaTT VARCHAR(10) PRIMARY KEY,
    MaLich VARCHAR(10) NOT NULL UNIQUE,
    ThoiGianThanhToan DATETIME NOT NULL,
    SoTien DECIMAL(12,2) NOT NULL,
    PhuongThuc NVARCHAR(50) NOT NULL,
    TrangThaiTT NVARCHAR(50) NOT NULL DEFAULT 'Đã thanh toán',
    CONSTRAINT FK_TT_LH FOREIGN KEY (MaLich) REFERENCES LICH_HEN(MaLich) ON UPDATE CASCADE ON DELETE RESTRICT,
    CONSTRAINT CHK_TT_SoTien CHECK (SoTien >= 0),
    CONSTRAINT CHK_TT_PhuongThuc CHECK (PhuongThuc IN ('Tiền mặt', 'Chuyển khoản', 'Thẻ')),
    CONSTRAINT CHK_TT_TrangThai CHECK (TrangThaiTT IN ('Chưa thanh toán', 'Đã thanh toán'))
);

-- =================================================================================
-- DỮ LIỆU MẪU (SAMPLE DATA)
-- =================================================================================

INSERT INTO KHACH_HANG (MaKH, HoTen, SDT, Email, DiaChi) VALUES
('KH001', 'Nguyễn Minh Tuấn', '0901234567', 'tuan.nguyen@gmail.com', '125 Cầu Giấy, Hà Nội'),
('KH002', 'Trần Thị Thu Hà', '0912345678', 'thuha.tran@gmail.com', '42 Nguyễn Chí Thanh, Đống Đa, Hà Nội'),
('KH003', 'Lê Hoàng Nam', '0987654321', 'hoangnam.le@outlook.com', '88 Lê Văn Lương, Thanh Xuân, Hà Nội'),
('KH004', 'Phạm Phương Thảo', '0934567890', 'phuongthao.pham@gmail.com', '19 Đội Cấn, Ba Đình, Hà Nội'),
('KH005', 'Vũ Đức Anh', '0978112233', 'ducanh.vu@gmail.com', '302 Minh Khai, Hai Bà Trưng, Hà Nội'),
('KH006', 'Hoàng Bích Ngọc', '0945667788', 'bichngoc.hoang@gmail.com', '56 Xuân Thủy, Cầu Giấy, Hà Nội');

INSERT INTO THU_CUNG (MaThuCung, MaKH, TenThuCung, Loai, Giong, GioiTinh, Tuoi, CanNang, GhiChu) VALUES
('TC001', 'KH001', 'Milo', 'Chó', 'Poodle Tiny', 'Đực', 18, 4.2, 'Lông xoăn dày, hơi nhát người lạ'),
('TC002', 'KH001', 'Miu', 'Mèo', 'Anh lông ngắn (ALN)', 'Cái', 12, 3.8, 'Hiền lành, thích gãi cằm'),
('TC003', 'KH002', 'Bơ', 'Chó', 'Corgi Pembroke', 'Đực', 24, 11.5, 'Đang thay lông mùa thu, dị ứng sữa tắm mùi mạnh'),
('TC004', 'KH003', 'Simba', 'Mèo', 'Maine Coon', 'Đực', 30, 7.4, 'Lông dài cần chải kỹ vùng bụng'),
('TC005', 'KH004', 'Lu', 'Chó', 'Golden Retriever', 'Đực', 36, 28.0, 'Rất thân thiện, thích tắm nước ấm'),
('TC006', 'KH005', 'Na', 'Mèo', 'Ba Tư (Persian)', 'Cái', 15, 3.5, 'Mắt nhạy cảm, cần vệ sinh tuyến lệ nhẹ nhàng'),
('TC007', 'KH006', 'Đậu', 'Chó', 'Phốc Sóc (Pomeranian)', 'Cái', 10, 2.6, 'Hiếu động, cần sấy lông ở nhiệt độ mát');

INSERT INTO NHAN_VIEN (MaNV, HoTen, NgaySinh, GioiTinh, SDT, ChucVu, CaLamViec, Luong, TrangThai) VALUES
('NV001', 'Đỗ Thanh Hương', '1995-04-12', 'Nữ', '0961112221', 'Quản lý', 'Hành chính (08:00-17:00)', 18000000, 'Đang làm việc'),
('NV002', 'Nguyễn Văn Khoa', '1999-08-20', 'Nam', '0961112222', 'Nhân viên chăm sóc', 'Ca sáng (08:00-12:00)', 9500000, 'Đang làm việc'),
('NV003', 'Lê Thị Mai Anh', '2000-11-05', 'Nữ', '0961112223', 'Nhân viên chăm sóc', 'Ca chiều (13:00-17:00)', 9800000, 'Đang làm việc'),
('NV004', 'Trần Quốc Bảo', '1998-02-18', 'Nam', '0961112224', 'Nhân viên chăm sóc', 'Hành chính (08:00-17:00)', 11000000, 'Đang làm việc'),
('NV005', 'Phạm Ngọc Lan', '2001-07-25', 'Nữ', '0961112225', 'Lễ tân', 'Hành chính (08:00-17:00)', 8500000, 'Đang làm việc'),
('NV006', 'Hoàng Minh Đức', '1997-09-14', 'Nam', '0961112226', 'Nhân viên chăm sóc', 'Ca tối (17:00-21:00)', 9200000, 'Nghỉ phép');

INSERT INTO DICH_VU (MaDV, TenDV, DonGia, ThoiLuongDuKien, MoTaDichVu) VALUES
('DV001', 'Tắm vệ sinh tiêu chuẩn 7 bước', 200000, 60, 'Tắm khử mùi, vệ sinh tai, vắt tuyến hôi, cắt mài móng, sấy chải bông lông'),
('DV002', 'Cắt tỉa tạo kiểu Grooming trọn gói', 380000, 120, 'Bao gồm tắm vệ sinh 7 bước + cắt tỉa lông nghệ thuật theo form giống'),
('DV003', 'Spa ủ dưỡng lông Collagen & Khử khuẩn', 320000, 90, 'Tắm thảo dược, ủ dưỡng phục hồi lông xơ rối và massage thư giãn'),
('DV004', 'Vệ sinh răng miệng & Lấy cao răng sóng âm', 250000, 45, 'Làm sạch mảng bám cao răng không gây mê, xịt thơm miệng kháng khuẩn'),
('DV005', 'Khách sạn thú cưng phòng VIP (Theo ngày)', 300000, 1440, 'Lưu trú phòng điều hòa riêng, camera 24/7, 3 bữa dinh dưỡng và vận động'),
('DV006', 'Gỡ rối lông chuyên sâu & Cạo vệ sinh', 180000, 60, 'Xử lý lông vón cục, cạo vệ sinh bàn chân, bụng và vùng hậu môn');

INSERT INTO LICH_HEN (MaLich, MaThuCung, MaDV, MaNV, NgayGioHen, SoLuong, TrangThai, GhiChu) VALUES
('LH001', 'TC001', 'DV002', 'NV002', '2026-11-01 09:00:00', 1, 'Hoàn thành', 'Cắt kiểu Teddy mặt tròn, bé hợp tác tốt'),
('LH002', 'TC002', 'DV001', 'NV004', '2026-11-02 10:00:00', 1, 'Hoàn thành', 'Tắm sạch, đã cắt móng chân trước'),
('LH003', 'TC003', 'DV003', 'NV003', '2026-11-03 14:00:00', 1, 'Hoàn thành', 'Dùng sữa tắm dịu nhẹ cho da nhạy cảm'),
('LH004', 'TC004', 'DV006', 'NV004', '2026-11-04 09:30:00', 1, 'Hoàn thành', 'Gỡ rối lông bụng thành công, không cần cạo'),
('LH005', 'TC005', 'DV001', 'NV002', '2026-11-05 08:30:00', 1, 'Hoàn thành', 'Chó lớn 28kg, sấy kỹ 2 lớp lông'),
('LH006', 'TC006', 'DV004', 'NV003', '2026-11-06 15:00:00', 1, 'Hoàn thành', 'Lấy sạch mảng bám răng hàm, vệ sinh mắt'),
('LH007', 'TC001', 'DV005', 'NV004', '2026-11-07 08:00:00', 2, 'Hoàn thành', 'Gửi khách sạn VIP 2 ngày cuối tuần'),
('LH008', 'TC007', 'DV002', 'NV002', '2026-11-08 09:30:00', 1, 'Hoàn thành', 'Tỉa form Boo tròn, sấy mát'),
('LH009', 'TC003', 'DV001', 'NV003', '2026-11-09 14:30:00', 1, 'Đang thực hiện', 'Khách hẹn đón lúc 16h00'),
('LH010', 'TC005', 'DV003', 'NV004', '2026-11-10 10:00:00', 1, 'Đã xác nhận', 'Khách đặt lịch trước qua điện thoại'),
('LH011', 'TC002', 'DV002', 'NV002', '2026-11-08 11:00:00', 1, 'Đã hủy', 'Khách bận đột xuất xin dời sang tuần sau');

INSERT INTO THANH_TOAN (MaTT, MaLich, ThoiGianThanhToan, SoTien, PhuongThuc, TrangThaiTT) VALUES
('TT001', 'LH001', '2026-11-01 11:10:00', 380000, 'Chuyển khoản', 'Đã thanh toán'),
('TT002', 'LH002', '2026-11-02 11:05:00', 200000, 'Tiền mặt', 'Đã thanh toán'),
('TT003', 'LH003', '2026-11-03 15:40:00', 320000, 'Chuyển khoản', 'Đã thanh toán'),
('TT004', 'LH004', '2026-11-04 10:45:00', 180000, 'Thẻ', 'Đã thanh toán'),
('TT005', 'LH005', '2026-11-05 10:00:00', 200000, 'Chuyển khoản', 'Đã thanh toán'),
('TT006', 'LH006', '2026-11-06 16:00:00', 250000, 'Tiền mặt', 'Đã thanh toán'),
('TT007', 'LH007', '2026-11-09 09:00:00', 600000, 'Chuyển khoản', 'Đã thanh toán'),
('TT008', 'LH008', '2026-11-08 11:45:00', 380000, 'Thẻ', 'Đã thanh toán');

-- =================================================================================
-- CÁC CÂU LỆNH SQL SỬ DỤNG TRONG MỤC 5, 6, 7
-- =================================================================================

-- [MỤC 5 - MÀN HÌNH NHẬP 1: ĐĂNG KÝ HỒ SƠ THÚ CƯNG MỚI (BẢNG THU_CUNG)]
-- Bước 1: Kiểm tra mã thú cưng đã tồn tại hay chưa và khách hàng sở hữu có hợp lệ không
-- SELECT COUNT(*) FROM THU_CUNG WHERE MaThuCung = ?;
-- SELECT COUNT(*) FROM KHACH_HANG WHERE MaKH = ?;
-- Bước 2: Thêm mới bản ghi vào bảng THU_CUNG (Ràng buộc CHECK tự động kiểm tra Tuoi > 0, CanNang > 0, GioiTinh hợp lệ)
-- INSERT INTO THU_CUNG (MaThuCung, MaKH, TenThuCung, Loai, Giong, GioiTinh, Tuoi, CanNang, GhiChu)
-- VALUES ('TC008', 'KH002', 'Kem', 'Mèo', 'Ragdoll', 'Cái', 14, 4.1, 'Lông dài mượt, đã tiêm đủ mũi');

-- [MỤC 5 - MÀN HÌNH NHẬP 2: TIẾP NHẬN & ĐẶT LỊCH HẸN CHĂM SÓC (BẢNG LICH_HEN)]
-- Bước 1: Kiểm tra nhân viên có đang làm việc và không bị trùng lịch trong khung giờ hẹn
-- SELECT TrangThai, CaLamViec FROM NHAN_VIEN WHERE MaNV = ?;
-- SELECT COUNT(*) FROM LICH_HEN
-- WHERE MaNV = ? AND TrangThai NOT IN ('Đã hủy', 'Không đến')
--   AND ABS( CAST((JULIANDAY(NgayGioHen) - JULIANDAY(?)) * 1440 AS INTEGER) ) < 60;
-- Bước 2: Thêm lịch hẹn mới
-- INSERT INTO LICH_HEN (MaLich, MaThuCung, MaDV, MaNV, NgayGioHen, SoLuong, TrangThai, GhiChu)
-- VALUES ('LH012', 'TC001', 'DV001', 'NV004', '2026-11-12 09:00:00', 1, 'Đã xác nhận', 'Tắm định kỳ hàng tuần');

-- [MỤC 6 - MÀN HÌNH TÌM KIẾM ĐA TIÊU CHÍ LỊCH HẸN & HỒ SƠ CHĂM SÓC]
SELECT 
    lh.MaLich,
    lh.NgayGioHen,
    kh.MaKH,
    kh.HoTen AS TenKhachHang,
    kh.SDT AS SDTKhachHang,
    tc.MaThuCung,
    tc.TenThuCung,
    tc.Loai,
    tc.Giong,
    tc.CanNang,
    dv.MaDV,
    dv.TenDV,
    dv.DonGia,
    lh.SoLuong,
    (dv.DonGia * lh.SoLuong) AS ThanhTienDuKien,
    nv.MaNV,
    nv.HoTen AS TenNhanVien,
    nv.CaLamViec,
    lh.TrangThai,
    COALESCE(tt.TrangThaiTT, 'Chưa thanh toán') AS TrangThaiThanhToan,
    lh.GhiChu
FROM LICH_HEN lh
JOIN THU_CUNG tc ON lh.MaThuCung = tc.MaThuCung
JOIN KHACH_HANG kh ON tc.MaKH = kh.MaKH
JOIN DICH_VU dv ON lh.MaDV = dv.MaDV
JOIN NHAN_VIEN nv ON lh.MaNV = nv.MaNV
LEFT JOIN THANH_TOAN tt ON lh.MaLich = tt.MaLich
WHERE 1=1
ORDER BY lh.NgayGioHen DESC;

-- [MỤC 7 - BÁO CÁO 1: BÁO CÁO DOANH THU DỊCH VỤ THEO KỲ & PHƯƠNG THỨC THANH TOÁN]
-- Người lập: Kế toán viên / Lễ tân trưởng -> Người nhận: Chủ cơ sở / Giám đốc cửa hàng
SELECT 
    dv.MaDV,
    dv.TenDV,
    dv.DonGia,
    COUNT(tt.MaTT) AS SoGiaoDichThanhToan,
    SUM(lh.SoLuong) AS TongSoLuongSuDung,
    SUM(dv.DonGia * lh.SoLuong) AS TongThanhTien,
    SUM(tt.SoTien) AS TongDoanhThuThucThu,
    SUM(CASE WHEN tt.PhuongThuc = 'Chuyển khoản' THEN tt.SoTien ELSE 0 END) AS ThuChuyenKhoan,
    SUM(CASE WHEN tt.PhuongThuc = 'Tiền mặt' THEN tt.SoTien ELSE 0 END) AS ThuTienMat,
    SUM(CASE WHEN tt.PhuongThuc = 'Thẻ' THEN tt.SoTien ELSE 0 END) AS ThuQuaThe
FROM DICH_VU dv
LEFT JOIN LICH_HEN lh ON dv.MaDV = lh.MaDV AND lh.TrangThai = 'Hoàn thành'
LEFT JOIN THANH_TOAN tt ON lh.MaLich = tt.MaLich AND tt.TrangThaiTT = 'Đã thanh toán'
GROUP BY dv.MaDV, dv.TenDV, dv.DonGia
ORDER BY TongDoanhThuThucThu DESC;

-- [MỤC 7 - BÁO CÁO 2: BÁO CÁO HIỆU SUẤT LÀM VIỆC CỦA NHÂN VIÊN & TÌNH HÌNH LỊCH HẸN]
-- Người lập: Quản lý cơ sở -> Người nhận: Bộ phận Nhân sự - Tiền lương & Ban Giám đốc
SELECT 
    nv.MaNV,
    nv.HoTen AS TenNhanVien,
    nv.ChucVu,
    nv.CaLamViec,
    nv.TrangThai AS TrangThaiLamViec,
    COUNT(lh.MaLich) AS TongSoLichPhanCong,
    SUM(CASE WHEN lh.TrangThai = 'Hoàn thành' THEN 1 ELSE 0 END) AS SoLichHoanThanh,
    SUM(CASE WHEN lh.TrangThai IN ('Đang thực hiện', 'Đã xác nhận', 'Mới tạo') THEN 1 ELSE 0 END) AS SoLichDangXuLy,
    SUM(CASE WHEN lh.TrangThai IN ('Đã hủy', 'Không đến') THEN 1 ELSE 0 END) AS SoLichHuy,
    COALESCE(SUM(CASE WHEN lh.TrangThai = 'Hoàn thành' THEN dv.ThoiLuongDuKien * lh.SoLuong ELSE 0 END), 0) AS TongSoPhutPhucVu,
    COALESCE(SUM(CASE WHEN tt.TrangThaiTT = 'Đã thanh toán' THEN tt.SoTien ELSE 0 END), 0) AS DoanhThuDongGop
FROM NHAN_VIEN nv
LEFT JOIN LICH_HEN lh ON nv.MaNV = lh.MaNV
LEFT JOIN DICH_VU dv ON lh.MaDV = dv.MaDV
LEFT JOIN THANH_TOAN tt ON lh.MaLich = tt.MaLich
WHERE nv.ChucVu = 'Nhân viên chăm sóc'
GROUP BY nv.MaNV, nv.HoTen, nv.ChucVu, nv.CaLamViec, nv.TrangThai
ORDER BY DoanhThuDongGop DESC, SoLichHoanThanh DESC;
