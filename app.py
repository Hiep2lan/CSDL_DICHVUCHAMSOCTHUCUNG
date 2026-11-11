import http.server
import json
import sqlite3
import urllib.parse
import os
import io
import zipfile
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "petcare.db")
SQL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "csdl_cham_soc_thu_cung.sql")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def build_xlsx(sheet_title, headers, rows, meta_lines=None):
    """Build a valid .xlsx file in memory using standard library zipfile + XML."""
    def esc(val):
        s = "" if val is None else str(val)
        return (
            s.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
        )

    def col_letter(idx):
        res = ""
        idx += 1
        while idx > 0:
            idx, rem = divmod(idx - 1, 26)
            res = chr(65 + rem) + res
        return res

    sheet_rows_xml = []
    r_idx = 1
    if meta_lines:
        for line in meta_lines:
            style = "1" if r_idx in (1, 2, 4) else "0"
            sheet_rows_xml.append(
                f'<row r="{r_idx}"><c r="A{r_idx}" s="{style}" t="inlineStr"><is><t>{esc(line)}</t></is></c></row>'
            )
            r_idx += 1
        r_idx += 1

    # Header row
    h_cells = []
    for c_i, h in enumerate(headers):
        ref = f"{col_letter(c_i)}{r_idx}"
        h_cells.append(f'<c r="{ref}" s="2" t="inlineStr"><is><t>{esc(h)}</t></is></c>')
    sheet_rows_xml.append(f'<row r="{r_idx}">{"".join(h_cells)}</row>')
    r_idx += 1

    # Data rows
    for row in rows:
        d_cells = []
        for c_i, val in enumerate(row):
            ref = f"{col_letter(c_i)}{r_idx}"
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                d_cells.append(f'<c r="{ref}" s="3"><v>{val}</v></c>')
            else:
                d_cells.append(
                    f'<c r="{ref}" s="3" t="inlineStr"><is><t>{esc(val)}</t></is></c>'
                )
        sheet_rows_xml.append(f'<row r="{r_idx}">{"".join(d_cells)}</row>')
        r_idx += 1

    worksheet_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="12" width="22" customWidth="1"/>
  </cols>
  <sheetData>
    {"".join(sheet_rows_xml)}
  </sheetData>
</worksheet>"""

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Times New Roman"/></font>
    <font><b/><sz val="12"/><name val="Times New Roman"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Times New Roman"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E3A8A"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color auto="1"/></left>
      <right style="thin"><color auto="1"/></right>
      <top style="thin"><color auto="1"/></top>
      <bottom style="thin"><color auto="1"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
  </cellXfs>
</styleSheet>"""

    workbook_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="{esc(sheet_title[:31])}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""

    wb_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

    root_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", root_rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels_xml)
        zf.writestr("xl/styles.xml", styles_xml)
        zf.writestr("xl/worksheets/sheet1.xml", worksheet_xml)
    return buf.getvalue()


class PetCareHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def send_json(self, data, status=200):
        raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        qs = urllib.parse.parse_qs(parsed.query)

        if path == "/api/init":
            conn = get_conn()
            kh = [dict(r) for r in conn.execute("SELECT * FROM KHACH_HANG ORDER BY MaKH").fetchall()]
            tc = [
                dict(r)
                for r in conn.execute(
                    """SELECT tc.*, kh.HoTen AS TenKH, kh.SDT AS SDT_KH
                       FROM THU_CUNG tc JOIN KHACH_HANG kh ON tc.MaKH = kh.MaKH
                       ORDER BY tc.MaThuCung"""
                ).fetchall()
            ]
            nv = [dict(r) for r in conn.execute("SELECT * FROM NHAN_VIEN ORDER BY MaNV").fetchall()]
            dv = [dict(r) for r in conn.execute("SELECT * FROM DICH_VU ORDER BY MaDV").fetchall()]
            lh = [
                dict(r)
                for r in conn.execute(
                    """SELECT lh.*, tc.TenThuCung, tc.Loai, tc.Giong, tc.MaKH,
                              kh.HoTen AS TenKhachHang, kh.SDT AS SDTKhachHang,
                              dv.TenDV, dv.DonGia, (dv.DonGia * lh.SoLuong) AS ThanhTienDuKien,
                              nv.HoTen AS TenNhanVien, nv.CaLamViec,
                              COALESCE(tt.TrangThaiTT, 'Chưa thanh toán') AS TrangThaiThanhToan
                       FROM LICH_HEN lh
                       JOIN THU_CUNG tc ON lh.MaThuCung = tc.MaThuCung
                       JOIN KHACH_HANG kh ON tc.MaKH = kh.MaKH
                       JOIN DICH_VU dv ON lh.MaDV = dv.MaDV
                       JOIN NHAN_VIEN nv ON lh.MaNV = nv.MaNV
                       LEFT JOIN THANH_TOAN tt ON lh.MaLich = tt.MaLich
                       ORDER BY lh.NgayGioHen DESC"""
                ).fetchall()
            ]
            tt = [dict(r) for r in conn.execute("SELECT * FROM THANH_TOAN ORDER BY MaTT").fetchall()]
            conn.close()
            return self.send_json({"khach_hang": kh, "thu_cung": tc, "nhan_vien": nv, "dich_vu": dv, "lich_hen": lh, "thanh_toan": tt})

        elif path == "/api/search":
            kw = qs.get("q", [""])[0].strip()
            loai = qs.get("loai", [""])[0].strip()
            madv = qs.get("madv", [""])[0].strip()
            manv = qs.get("manv", [""])[0].strip()
            trangthai = qs.get("trangthai", [""])[0].strip()
            tu_ngay = qs.get("tu_ngay", [""])[0].strip()
            den_ngay = qs.get("den_ngay", [""])[0].strip()

            sql = """SELECT 
                lh.MaLich, lh.NgayGioHen,
                kh.MaKH, kh.HoTen AS TenKhachHang, kh.SDT AS SDTKhachHang,
                tc.MaThuCung, tc.TenThuCung, tc.Loai, tc.Giong, tc.CanNang,
                dv.MaDV, dv.TenDV, dv.DonGia, lh.SoLuong,
                (dv.DonGia * lh.SoLuong) AS ThanhTienDuKien,
                nv.MaNV, nv.HoTen AS TenNhanVien, nv.CaLamViec,
                lh.TrangThai,
                COALESCE(tt.TrangThaiTT, 'Chưa thanh toán') AS TrangThaiThanhToan,
                lh.GhiChu
            FROM LICH_HEN lh
            JOIN THU_CUNG tc ON lh.MaThuCung = tc.MaThuCung
            JOIN KHACH_HANG kh ON tc.MaKH = kh.MaKH
            JOIN DICH_VU dv ON lh.MaDV = dv.MaDV
            JOIN NHAN_VIEN nv ON lh.MaNV = nv.MaNV
            LEFT JOIN THANH_TOAN tt ON lh.MaLich = tt.MaLich
            WHERE 1=1"""
            params = []
            if kw:
                sql += " AND (lh.MaLich LIKE ? OR kh.HoTen LIKE ? OR kh.SDT LIKE ? OR tc.TenThuCung LIKE ?)"
                like_kw = f"%{kw}%"
                params.extend([like_kw, like_kw, like_kw, like_kw])
            if loai:
                sql += " AND tc.Loai = ?"
                params.append(loai)
            if madv:
                sql += " AND lh.MaDV = ?"
                params.append(madv)
            if manv:
                sql += " AND lh.MaNV = ?"
                params.append(manv)
            if trangthai:
                sql += " AND lh.TrangThai = ?"
                params.append(trangthai)
            if tu_ngay:
                sql += " AND DATE(lh.NgayGioHen) >= DATE(?)"
                params.append(tu_ngay)
            if den_ngay:
                sql += " AND DATE(lh.NgayGioHen) <= DATE(?)"
                params.append(den_ngay)
            sql += " ORDER BY lh.NgayGioHen DESC"

            conn = get_conn()
            rows = [dict(r) for r in conn.execute(sql, params).fetchall()]
            conn.close()
            return self.send_json({"results": rows, "sql": sql, "params": params})

        elif path == "/api/reports":
            conn = get_conn()
            rep1_sql = """SELECT 
                dv.MaDV,
                dv.TenDV,
                dv.DonGia,
                COUNT(tt.MaTT) AS SoGiaoDichThanhToan,
                COALESCE(SUM(CASE WHEN tt.MaTT IS NOT NULL THEN lh.SoLuong ELSE 0 END), 0) AS TongSoLuongSuDung,
                COALESCE(SUM(CASE WHEN tt.MaTT IS NOT NULL THEN dv.DonGia * lh.SoLuong ELSE 0 END), 0) AS TongThanhTien,
                COALESCE(SUM(tt.SoTien), 0) AS TongDoanhThuThucThu,
                COALESCE(SUM(CASE WHEN tt.PhuongThuc = 'Chuyển khoản' THEN tt.SoTien ELSE 0 END), 0) AS ThuChuyenKhoan,
                COALESCE(SUM(CASE WHEN tt.PhuongThuc = 'Tiền mặt' THEN tt.SoTien ELSE 0 END), 0) AS ThuTienMat,
                COALESCE(SUM(CASE WHEN tt.PhuongThuc = 'Thẻ' THEN tt.SoTien ELSE 0 END), 0) AS ThuQuaThe
            FROM DICH_VU dv
            LEFT JOIN LICH_HEN lh ON dv.MaDV = lh.MaDV AND lh.TrangThai = 'Hoàn thành'
            LEFT JOIN THANH_TOAN tt ON lh.MaLich = tt.MaLich AND tt.TrangThaiTT = 'Đã thanh toán'
            GROUP BY dv.MaDV, dv.TenDV, dv.DonGia
            ORDER BY TongDoanhThuThucThu DESC"""
            rep1 = [dict(r) for r in conn.execute(rep1_sql).fetchall()]

            rep2_sql = """SELECT 
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
            ORDER BY DoanhThuDongGop DESC, SoLichHoanThanh DESC"""
            rep2 = [dict(r) for r in conn.execute(rep2_sql).fetchall()]
            conn.close()
            return self.send_json({"report1": rep1, "report2": rep2, "sql1": rep1_sql, "sql2": rep2_sql})

        elif path == "/api/export_xlsx":
            rep_type = qs.get("type", ["1"])[0]
            conn = get_conn()
            if rep_type == "1":
                rows_db = conn.execute(
                    """SELECT dv.MaDV, dv.TenDV, dv.DonGia,
                              COUNT(tt.MaTT) AS SoGD,
                              COALESCE(SUM(CASE WHEN tt.MaTT IS NOT NULL THEN lh.SoLuong ELSE 0 END), 0) AS TongSL,
                              COALESCE(SUM(CASE WHEN tt.PhuongThuc = 'Tiền mặt' THEN tt.SoTien ELSE 0 END), 0) AS TienMat,
                              COALESCE(SUM(CASE WHEN tt.PhuongThuc = 'Chuyển khoản' THEN tt.SoTien ELSE 0 END), 0) AS ChuyenKhoan,
                              COALESCE(SUM(CASE WHEN tt.PhuongThuc = 'Thẻ' THEN tt.SoTien ELSE 0 END), 0) AS QuaThe,
                              COALESCE(SUM(tt.SoTien), 0) AS TongThu
                       FROM DICH_VU dv
                       LEFT JOIN LICH_HEN lh ON dv.MaDV = lh.MaDV AND lh.TrangThai = 'Hoàn thành'
                       LEFT JOIN THANH_TOAN tt ON lh.MaLich = tt.MaLich AND tt.TrangThaiTT = 'Đã thanh toán'
                       GROUP BY dv.MaDV, dv.TenDV, dv.DonGia
                       ORDER BY TongThu DESC"""
                ).fetchall()
                meta = [
                    "TRUNG TÂM CHĂM SÓC THÚ CƯNG PETCARE SPA & CLINIC",
                    "BÁO CÁO TỔNG HỢP DOANH THU DỊCH VỤ CHĂM SÓC THÚ CƯNG (MẪU BC-01/TC-KT)",
                    "Người lập báo cáo: Phạm Ngọc Lan (Nhân viên Lễ tân - Thu ngân) | Người nhận: Đỗ Thanh Hương (Quản lý cơ sở / Chủ cửa hàng)",
                    f"Ngày xuất báo cáo: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                ]
                headers = [
                    "Mã DV",
                    "Tên Dịch Vụ",
                    "Đơn Giá (VNĐ)",
                    "Số Giao Dịch",
                    "Tổng Số Lượng",
                    "Thu Tiền Mặt (VNĐ)",
                    "Thu Chuyển Khoản (VNĐ)",
                    "Thu Thẻ (VNĐ)",
                    "Tổng Thực Thu (VNĐ)",
                ]
                data_rows = [list(r) for r in rows_db]
                xlsx_bytes = build_xlsx("BaoCaoDoanhThu", headers, data_rows, meta)
                fname = "Bao_Cao_Doanh_Thu_Dich_Vu_PetCare.xlsx"
            else:
                rows_db = conn.execute(
                    """SELECT nv.MaNV, nv.HoTen, nv.CaLamViec, nv.TrangThai,
                              COUNT(lh.MaLich) AS TongLich,
                              SUM(CASE WHEN lh.TrangThai = 'Hoàn thành' THEN 1 ELSE 0 END) AS HoanThanh,
                              SUM(CASE WHEN lh.TrangThai IN ('Đang thực hiện', 'Đã xác nhận', 'Mới tạo') THEN 1 ELSE 0 END) AS DangXuLy,
                              SUM(CASE WHEN lh.TrangThai IN ('Đã hủy', 'Không đến') THEN 1 ELSE 0 END) AS DaHuy,
                              COALESCE(SUM(CASE WHEN lh.TrangThai = 'Hoàn thành' THEN dv.ThoiLuongDuKien * lh.SoLuong ELSE 0 END), 0) AS SoPhut,
                              COALESCE(SUM(CASE WHEN tt.TrangThaiTT = 'Đã thanh toán' THEN tt.SoTien ELSE 0 END), 0) AS DoanhThu
                       FROM NHAN_VIEN nv
                       LEFT JOIN LICH_HEN lh ON nv.MaNV = lh.MaNV
                       LEFT JOIN DICH_VU dv ON lh.MaDV = dv.MaDV
                       LEFT JOIN THANH_TOAN tt ON lh.MaLich = tt.MaLich
                       WHERE nv.ChucVu = 'Nhân viên chăm sóc'
                       GROUP BY nv.MaNV, nv.HoTen, nv.CaLamViec, nv.TrangThai
                       ORDER BY DoanhThu DESC"""
                ).fetchall()
                meta = [
                    "TRUNG TÂM CHĂM SÓC THÚ CƯNG PETCARE SPA & CLINIC",
                    "BÁO CÁO HIỆU SUẤT LÀM VIỆC CỦA NHÂN VIÊN & TÌNH HÌNH THỰC HIỆN LỊCH HẸN (MẪU BC-02/NS-ĐH)",
                    "Người lập báo cáo: Đỗ Thanh Hương (Quản lý cơ sở) | Người nhận: Bộ phận Nhân sự - Tiền lương & Ban Giám đốc",
                    f"Ngày xuất báo cáo: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                ]
                headers = [
                    "Mã NV",
                    "Họ Tên Nhân Viên",
                    "Ca Làm Việc",
                    "Trạng Thái",
                    "Tổng Lịch Phân Công",
                    "Lịch Hoàn Thành",
                    "Lịch Đang Xử Lý",
                    "Lịch Đã Hủy",
                    "Tổng Phút Phục Vụ",
                    "Doanh Thu Đóng Góp (VNĐ)",
                ]
                data_rows = [list(r) for r in rows_db]
                xlsx_bytes = build_xlsx("BaoCaoHieuSuatNV", headers, data_rows, meta)
                fname = "Bao_Cao_Hieu_Suat_Nhan_Vien_PetCare.xlsx"
            conn.close()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(xlsx_bytes)))
            self.end_headers()
            self.wfile.write(xlsx_bytes)
            return

        return super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        length = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(length).decode("utf-8")) if length > 0 else {}

        if parsed.path == "/api/khach_hang":
            errors = []
            makh = str(body.get("MaKH", "")).strip()
            hoten = str(body.get("HoTen", "")).strip()
            sdt = str(body.get("SDT", "")).strip()
            email = str(body.get("Email", "")).strip() or None
            diachi = str(body.get("DiaChi", "")).strip()

            if not makh or not makh.startswith("KH") or len(makh) < 4:
                errors.append("Mã khách hàng (MaKH) không hợp lệ: Phải bắt đầu bằng 'KH' kèm số (VD: KH007).")
            if not hoten:
                errors.append("Ràng buộc NOT NULL: Họ tên khách hàng (HoTen) không được để trống.")
            if not sdt or not (10 <= len(sdt) <= 11) or not sdt.startswith("0") or not sdt.isdigit():
                errors.append(f"Vi phạm ràng buộc CHECK (CHK_KH_SDT): Số điện thoại '{sdt}' phải gồm 10-11 chữ số và bắt đầu bằng số 0.")

            conn = get_conn()
            if makh and conn.execute("SELECT 1 FROM KHACH_HANG WHERE MaKH = ?", (makh,)).fetchone():
                errors.append(f"Vi phạm Khóa chính (PRIMARY KEY): Mã khách hàng '{makh}' đã tồn tại trong bảng KHACH_HANG.")
            if sdt and conn.execute("SELECT 1 FROM KHACH_HANG WHERE SDT = ?", (sdt,)).fetchone():
                errors.append(f"Vi phạm ràng buộc Duy nhất (UNIQUE): Số điện thoại '{sdt}' đã được đăng ký cho khách hàng khác.")
            if email and conn.execute("SELECT 1 FROM KHACH_HANG WHERE Email = ?", (email,)).fetchone():
                errors.append(f"Vi phạm ràng buộc Duy nhất (UNIQUE): Email '{email}' đã tồn tại trong bảng KHACH_HANG.")

            if errors:
                conn.close()
                return self.send_json({"ok": False, "errors": errors}, status=400)

            try:
                conn.execute(
                    "INSERT INTO KHACH_HANG (MaKH, HoTen, SDT, Email, DiaChi) VALUES (?, ?, ?, ?, ?)",
                    (makh, hoten, sdt, email, diachi),
                )
                conn.commit()
            except sqlite3.IntegrityError as e:
                conn.close()
                return self.send_json({"ok": False, "errors": [f"Lỗi toàn vẹn CSDL (SQLite IntegrityError): {str(e)}"]}, status=400)
            conn.close()
            email_sql = f"'{email}'" if email else "NULL"
            return self.send_json(
                {
                    "ok": True,
                    "message": f"Đã thêm khách hàng mới '{hoten}' ({makh}) thành công vào bảng KHACH_HANG!",
                    "sql": f"INSERT INTO KHACH_HANG (MaKH, HoTen, SDT, Email, DiaChi) VALUES ('{makh}', '{hoten}', '{sdt}', {email_sql}, '{diachi}');",
                }
            )

        elif parsed.path == "/api/thu_cung":
            # Validate fields
            errors = []
            matc = str(body.get("MaThuCung", "")).strip()
            makh = str(body.get("MaKH", "")).strip()
            tentc = str(body.get("TenThuCung", "")).strip()
            loai = str(body.get("Loai", "")).strip()
            giong = str(body.get("Giong", "")).strip()
            gioitinh = str(body.get("GioiTinh", "")).strip()
            ghichu = str(body.get("GhiChu", "")).strip()
            new_kh = body.get("new_customer")

            conn = get_conn()

            if new_kh and isinstance(new_kh, dict):
                makh = str(new_kh.get("MaKH", "")).strip()
                kh_hoten = str(new_kh.get("HoTen", "")).strip()
                kh_sdt = str(new_kh.get("SDT", "")).strip()
                kh_email = str(new_kh.get("Email", "")).strip() or None
                kh_diachi = str(new_kh.get("DiaChi", "")).strip()

                if not makh or not makh.startswith("KH") or len(makh) < 4:
                    errors.append("Mã khách hàng mới (MaKH) không hợp lệ: Phải bắt đầu bằng 'KH' kèm số (VD: KH007).")
                if not kh_hoten:
                    errors.append("Ràng buộc NOT NULL: Họ tên khách hàng mới (HoTen) không được để trống.")
                if not kh_sdt or not (10 <= len(kh_sdt) <= 11) or not kh_sdt.startswith("0") or not kh_sdt.isdigit():
                    errors.append(f"Vi phạm ràng buộc CHECK (CHK_KH_SDT): SĐT khách hàng '{kh_sdt}' phải gồm 10-11 chữ số và bắt đầu bằng số 0.")
                if makh and conn.execute("SELECT 1 FROM KHACH_HANG WHERE MaKH = ?", (makh,)).fetchone():
                    errors.append(f"Vi phạm Khóa chính (PRIMARY KEY): Mã khách hàng '{makh}' đã tồn tại trong bảng KHACH_HANG.")
                if kh_sdt and conn.execute("SELECT 1 FROM KHACH_HANG WHERE SDT = ?", (kh_sdt,)).fetchone():
                    errors.append(f"Vi phạm ràng buộc UNIQUE: Số điện thoại '{kh_sdt}' đã tồn tại trong bảng KHACH_HANG.")
                if kh_email and conn.execute("SELECT 1 FROM KHACH_HANG WHERE Email = ?", (kh_email,)).fetchone():
                    errors.append(f"Vi phạm ràng buộc UNIQUE: Email '{kh_email}' đã tồn tại trong bảng KHACH_HANG.")

            if not matc or not matc.startswith("TC") or len(matc) < 4:
                errors.append("Mã thú cưng (MaThuCung) không hợp lệ: Phải bắt đầu bằng 'TC' kèm số (VD: TC008).")
            if not makh:
                errors.append("Chủ sở hữu (MaKH) không được để trống.")
            if not tentc:
                errors.append("Tên thú cưng (TenThuCung) không được để trống.")
            if not loai:
                errors.append("Loài thú cưng (Loai) không được để trống.")
            if gioitinh not in ("Đực", "Cái", "Chưa xác định"):
                errors.append("Giới tính thú cưng chỉ nhận giá trị: 'Đực', 'Cái' hoặc 'Chưa xác định'.")

            try:
                tuoi = int(body.get("Tuoi", 0))
                if tuoi <= 0 or tuoi > 360:
                    errors.append(f"Tuổi thú cưng ({tuoi} tháng) vi phạm ràng buộc CHECK (0 < Tuoi <= 360 tháng).")
            except ValueError:
                errors.append("Tuổi thú cưng phải là số nguyên hợp lệ.")
                tuoi = 0

            try:
                cannang = float(body.get("CanNang", 0))
                if cannang <= 0 or cannang > 150:
                    errors.append(f"Cân nặng ({cannang} kg) vi phạm ràng buộc CHECK (0 < CanNang <= 150 kg).")
            except ValueError:
                errors.append("Cân nặng phải là số thực hợp lệ.")
                cannang = 0

            if matc:
                dup = conn.execute("SELECT 1 FROM THU_CUNG WHERE MaThuCung = ?", (matc,)).fetchone()
                if dup:
                    errors.append(f"Vi phạm Khóa chính (PRIMARY KEY): Mã thú cưng '{matc}' đã tồn tại trong bảng THU_CUNG.")
            if makh and not new_kh:
                kh_ok = conn.execute("SELECT 1 FROM KHACH_HANG WHERE MaKH = ?", (makh,)).fetchone()
                if not kh_ok:
                    errors.append(f"Vi phạm Khóa ngoại (FOREIGN KEY): Mã khách hàng '{makh}' không tồn tại trong bảng KHACH_HANG.")

            if errors:
                conn.close()
                return self.send_json({"ok": False, "errors": errors}, status=400)

            try:
                sql_statements = []
                if new_kh and isinstance(new_kh, dict):
                    conn.execute(
                        "INSERT INTO KHACH_HANG (MaKH, HoTen, SDT, Email, DiaChi) VALUES (?, ?, ?, ?, ?)",
                        (makh, kh_hoten, kh_sdt, kh_email, kh_diachi),
                    )
                    email_sql = f"'{kh_email}'" if kh_email else "NULL"
                    sql_statements.append(
                        f"-- Bước 1: Thêm Khách hàng mới vào bảng cha KHACH_HANG\nINSERT INTO KHACH_HANG (MaKH, HoTen, SDT, Email, DiaChi) VALUES ('{makh}', '{kh_hoten}', '{kh_sdt}', {email_sql}, '{kh_diachi}');"
                    )
                sql = """INSERT INTO THU_CUNG (MaThuCung, MaKH, TenThuCung, Loai, Giong, GioiTinh, Tuoi, CanNang, GhiChu)
                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"""
                conn.execute(sql, (matc, makh, tentc, loai, giong, gioitinh, tuoi, cannang, ghichu))
                sql_statements.append(
                    f"-- Bước 2: Thêm Hồ sơ Thú cưng vào bảng con THU_CUNG (tham chiếu MaKH = '{makh}')\nINSERT INTO THU_CUNG (MaThuCung, MaKH, TenThuCung, Loai, Giong, GioiTinh, Tuoi, CanNang, GhiChu) VALUES ('{matc}', '{makh}', '{tentc}', '{loai}', '{giong}', '{gioitinh}', {tuoi}, {cannang}, '{ghichu}');"
                )
                conn.commit()
            except sqlite3.IntegrityError as e:
                conn.close()
                return self.send_json({"ok": False, "errors": [f"Lỗi toàn vẹn CSDL (SQLite IntegrityError): {str(e)}"]}, status=400)
            conn.close()
            return self.send_json(
                {
                    "ok": True,
                    "message": f"Đã lưu thành công hồ sơ thú cưng '{tentc}' ({matc}) vào CSDL!",
                    "sql": "\n".join(sql_statements),
                }
            )

        elif parsed.path == "/api/lich_hen":
            errors = []
            malich = str(body.get("MaLich", "")).strip()
            matc = str(body.get("MaThuCung", "")).strip()
            madv = str(body.get("MaDV", "")).strip()
            manv = str(body.get("MaNV", "")).strip()
            ngaygio = str(body.get("NgayGioHen", "")).strip().replace("T", " ")
            if len(ngaygio) == 16:
                ngaygio += ":00"
            trangthai = str(body.get("TrangThai", "Đã xác nhận")).strip()
            ghichu = str(body.get("GhiChu", "")).strip()

            if not malich or not malich.startswith("LH") or len(malich) < 4:
                errors.append("Mã lịch hẹn (MaLich) không hợp lệ: Phải bắt đầu bằng 'LH' kèm số (VD: LH012).")
            if not matc:
                errors.append("Phải chọn Thú cưng (MaThuCung).")
            if not madv:
                errors.append("Phải chọn Dịch vụ (MaDV).")
            if not manv:
                errors.append("Phải chọn Nhân viên thực hiện (MaNV).")
            if not ngaygio:
                errors.append("Ngày giờ hẹn (NgayGioHen) không được để trống.")

            try:
                soluong = int(body.get("SoLuong", 0))
                if soluong <= 0:
                    errors.append(f"Số lượng dịch vụ ({soluong}) vi phạm ràng buộc CHECK (SoLuong > 0).")
            except ValueError:
                errors.append("Số lượng dịch vụ phải là số nguyên lớn hơn 0.")
                soluong = 0

            conn = get_conn()
            if malich:
                if conn.execute("SELECT 1 FROM LICH_HEN WHERE MaLich = ?", (malich,)).fetchone():
                    errors.append(f"Vi phạm Khóa chính (PRIMARY KEY): Mã lịch hẹn '{malich}' đã tồn tại trong bảng LICH_HEN.")
            if matc:
                if not conn.execute("SELECT 1 FROM THU_CUNG WHERE MaThuCung = ?", (matc,)).fetchone():
                    errors.append(f"Vi phạm Khóa ngoại: Thú cưng '{matc}' không tồn tại.")
            if madv:
                if not conn.execute("SELECT 1 FROM DICH_VU WHERE MaDV = ?", (madv,)).fetchone():
                    errors.append(f"Vi phạm Khóa ngoại: Dịch vụ '{madv}' không tồn tại.")

            if manv:
                nv_row = conn.execute("SELECT HoTen, TrangThai, CaLamViec, ChucVu FROM NHAN_VIEN WHERE MaNV = ?", (manv,)).fetchone()
                if not nv_row:
                    errors.append(f"Vi phạm Khóa ngoại: Nhân viên '{manv}' không tồn tại.")
                else:
                    if nv_row["TrangThai"] != "Đang làm việc":
                        errors.append(
                            f"Vi phạm quy tắc nghiệp vụ: Nhân viên '{nv_row['HoTen']}' ({manv}) hiện đang ở trạng thái '{nv_row['TrangThai']}', không thể nhận lịch hẹn."
                        )
                    # Check shift match
                    if ngaygio and len(ngaygio) >= 16:
                        try:
                            dt = datetime.strptime(ngaygio[:16], "%Y-%m-%d %H:%M")
                            hr = dt.hour + dt.minute / 60.0
                            ca = nv_row["CaLamViec"]
                            if "Ca sáng" in ca and not (8 <= hr <= 12):
                                errors.append(
                                    f"Vi phạm quy tắc phân công ca: Nhân viên '{nv_row['HoTen']}' làm '{ca}' nhưng giờ hẹn là {ngaygio[11:16]}."
                                )
                            elif "Ca chiều" in ca and not (13 <= hr <= 17):
                                errors.append(
                                    f"Vi phạm quy tắc phân công ca: Nhân viên '{nv_row['HoTen']}' làm '{ca}' nhưng giờ hẹn là {ngaygio[11:16]}."
                                )
                            elif "Ca tối" in ca and not (17 <= hr <= 21):
                                errors.append(
                                    f"Vi phạm quy tắc phân công ca: Nhân viên '{nv_row['HoTen']}' làm '{ca}' nhưng giờ hẹn là {ngaygio[11:16]}."
                                )
                            elif "Hành chính" in ca and not (8 <= hr <= 17):
                                errors.append(
                                    f"Vi phạm quy tắc phân công ca: Nhân viên '{nv_row['HoTen']}' làm '{ca}' nhưng giờ hẹn là {ngaygio[11:16]}."
                                )
                        except ValueError:
                            errors.append("Định dạng Ngày giờ hẹn không hợp lệ (YYYY-MM-DD HH:MM:SS).")

                    # Check overlapping appointment for this staff within 60 mins
                    if ngaygio:
                        overlap = conn.execute(
                            """SELECT MaLich, NgayGioHen FROM LICH_HEN
                               WHERE MaNV = ? AND TrangThai NOT IN ('Đã hủy', 'Không đến')
                                 AND ABS((JULIANDAY(NgayGioHen) - JULIANDAY(?)) * 1440) < 60""",
                            (manv, ngaygio),
                        ).fetchone()
                        if overlap:
                            errors.append(
                                f"Vi phạm quy tắc chống trùng lịch: Nhân viên '{nv_row['HoTen']}' đã có lịch hẹn '{overlap['MaLich']}' lúc {overlap['NgayGioHen']} (chồng lấn thời gian thực hiện)."
                            )

            if errors:
                conn.close()
                return self.send_json({"ok": False, "errors": errors}, status=400)

            try:
                conn.execute(
                    """INSERT INTO LICH_HEN (MaLich, MaThuCung, MaDV, MaNV, NgayGioHen, SoLuong, TrangThai, GhiChu)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                    (malich, matc, madv, manv, ngaygio, soluong, trangthai, ghichu),
                )
                conn.commit()
            except sqlite3.IntegrityError as e:
                conn.close()
                return self.send_json({"ok": False, "errors": [f"Lỗi toàn vẹn CSDL: {str(e)}"]}, status=400)

            conn.close()
            return self.send_json(
                {
                    "ok": True,
                    "message": f"Đã tiếp nhận và tạo lịch hẹn '{malich}' thành công vào bảng LICH_HEN!",
                    "sql": f"INSERT INTO LICH_HEN (MaLich, MaThuCung, MaDV, MaNV, NgayGioHen, SoLuong, TrangThai, GhiChu) VALUES ('{malich}', '{matc}', '{madv}', '{manv}', '{ngaygio}', {soluong}, '{trangthai}', '{ghichu}');",
                }
            )

        elif parsed.path == "/api/lich_hen/status":
            malich = str(body.get("MaLich", "")).strip()
            trangthai = str(body.get("TrangThai", "")).strip()
            phuongthuc = str(body.get("PhuongThuc", "Chuyển khoản")).strip()

            valid_statuses = ("Mới tạo", "Đã xác nhận", "Đang thực hiện", "Hoàn thành", "Đã hủy", "Không đến")
            if not malich:
                return self.send_json({"ok": False, "errors": ["Mã lịch hẹn (MaLich) không được để trống."]}, status=400)
            if trangthai not in valid_statuses:
                return self.send_json(
                    {"ok": False, "errors": [f"Vi phạm ràng buộc CHECK (CHK_LH_TrangThai): Trạng thái '{trangthai}' không hợp lệ."]},
                    status=400,
                )

            conn = get_conn()
            lh_row = conn.execute(
                """SELECT lh.MaLich, lh.SoLuong, dv.DonGia
                   FROM LICH_HEN lh JOIN DICH_VU dv ON lh.MaDV = dv.MaDV
                   WHERE lh.MaLich = ?""",
                (malich,),
            ).fetchone()
            if not lh_row:
                conn.close()
                return self.send_json({"ok": False, "errors": [f"Không tìm thấy lịch hẹn '{malich}' trong CSDL."]}, status=404)

            sql_logs = []
            try:
                conn.execute("UPDATE LICH_HEN SET TrangThai = ? WHERE MaLich = ?", (trangthai, malich))
                sql_logs.append(f"UPDATE LICH_HEN SET TrangThai = '{trangthai}' WHERE MaLich = '{malich}';")

                new_tt = None
                if trangthai == "Hoàn thành":
                    existing_tt = conn.execute("SELECT MaTT FROM THANH_TOAN WHERE MaLich = ?", (malich,)).fetchone()
                    sotien = float(lh_row["DonGia"]) * int(lh_row["SoLuong"])
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    if existing_tt:
                        conn.execute(
                            "UPDATE THANH_TOAN SET TrangThaiTT = 'Đã thanh toán', PhuongThuc = ?, SoTien = ?, ThoiGianThanhToan = ? WHERE MaLich = ?",
                            (phuongthuc, sotien, now_str, malich),
                        )
                        sql_logs.append(
                            f"UPDATE THANH_TOAN SET TrangThaiTT = 'Đã thanh toán', PhuongThuc = '{phuongthuc}', SoTien = {int(sotien)} WHERE MaLich = '{malich}';"
                        )
                    else:
                        tt_rows = conn.execute("SELECT MaTT FROM THANH_TOAN").fetchall()
                        max_num = 0
                        for r in tt_rows:
                            digits = "".join(ch for ch in r["MaTT"] if ch.isdigit())
                            if digits:
                                max_num = max(max_num, int(digits))
                        matt = f"TT{max_num + 1:03d}"
                        conn.execute(
                            """INSERT INTO THANH_TOAN (MaTT, MaLich, ThoiGianThanhToan, SoTien, PhuongThuc, TrangThaiTT)
                               VALUES (?, ?, ?, ?, ?, 'Đã thanh toán')""",
                            (matt, malich, now_str, sotien, phuongthuc),
                        )
                        sql_logs.append(
                            f"-- Tự động ghi nhận phiếu thu vào bảng THANH_TOAN khi hoàn thành dịch vụ:\nINSERT INTO THANH_TOAN (MaTT, MaLich, ThoiGianThanhToan, SoTien, PhuongThuc, TrangThaiTT) VALUES ('{matt}', '{malich}', '{now_str}', {int(sotien)}, '{phuongthuc}', 'Đã thanh toán');"
                        )
                        new_tt = {
                            "MaTT": matt,
                            "MaLich": malich,
                            "ThoiGianThanhToan": now_str,
                            "SoTien": sotien,
                            "PhuongThuc": phuongthuc,
                            "TrangThaiTT": "Đã thanh toán",
                        }
                elif trangthai in ("Đã hủy", "Không đến"):
                    existing_tt = conn.execute("SELECT MaTT FROM THANH_TOAN WHERE MaLich = ?", (malich,)).fetchone()
                    if existing_tt:
                        conn.execute("UPDATE THANH_TOAN SET TrangThaiTT = 'Chưa thanh toán' WHERE MaLich = ?", (malich,))
                        sql_logs.append(f"UPDATE THANH_TOAN SET TrangThaiTT = 'Chưa thanh toán' WHERE MaLich = '{malich}';")

                conn.commit()
            except sqlite3.Error as e:
                conn.close()
                return self.send_json({"ok": False, "errors": [f"Lỗi cập nhật CSDL: {str(e)}"]}, status=400)

            conn.close()
            return self.send_json(
                {
                    "ok": True,
                    "message": f"Đã cập nhật trạng thái lịch hẹn '{malich}' sang '{trangthai}' thành công!",
                    "sql": "\n".join(sql_logs),
                    "thanh_toan": new_tt,
                }
            )

        elif parsed.path == "/api/thanh_toan/method":
            malich = str(body.get("MaLich", "")).strip()
            phuongthuc = str(body.get("PhuongThuc", "")).strip()

            if not malich:
                return self.send_json({"ok": False, "errors": ["Mã lịch hẹn (MaLich) không được để trống."]}, status=400)

            conn = get_conn()
            lh_row = conn.execute(
                """SELECT lh.MaLich, lh.SoLuong, dv.DonGia
                   FROM LICH_HEN lh JOIN DICH_VU dv ON lh.MaDV = dv.MaDV
                   WHERE lh.MaLich = ?""",
                (malich,),
            ).fetchone()
            if not lh_row:
                conn.close()
                return self.send_json({"ok": False, "errors": [f"Không tìm thấy lịch hẹn '{malich}'."]}, status=404)

            sql_logs = []
            new_tt = None
            try:
                existing_tt = conn.execute("SELECT MaTT FROM THANH_TOAN WHERE MaLich = ?", (malich,)).fetchone()
                if phuongthuc == "Chưa thanh toán":
                    if existing_tt:
                        conn.execute("UPDATE THANH_TOAN SET TrangThaiTT = 'Chưa thanh toán' WHERE MaLich = ?", (malich,))
                        sql_logs.append(f"UPDATE THANH_TOAN SET TrangThaiTT = 'Chưa thanh toán' WHERE MaLich = '{malich}';")
                elif phuongthuc in ("Tiền mặt", "Chuyển khoản", "Thẻ"):
                    sotien = float(lh_row["DonGia"]) * int(lh_row["SoLuong"])
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    if existing_tt:
                        conn.execute(
                            "UPDATE THANH_TOAN SET PhuongThuc = ?, TrangThaiTT = 'Đã thanh toán', SoTien = ?, ThoiGianThanhToan = ? WHERE MaLich = ?",
                            (phuongthuc, sotien, now_str, malich),
                        )
                        sql_logs.append(
                            f"UPDATE THANH_TOAN SET PhuongThuc = '{phuongthuc}', TrangThaiTT = 'Đã thanh toán', SoTien = {int(sotien)} WHERE MaLich = '{malich}';"
                        )
                    else:
                        tt_rows = conn.execute("SELECT MaTT FROM THANH_TOAN").fetchall()
                        max_num = 0
                        for r in tt_rows:
                            digits = "".join(ch for ch in r["MaTT"] if ch.isdigit())
                            if digits:
                                max_num = max(max_num, int(digits))
                        matt = f"TT{max_num + 1:03d}"
                        conn.execute(
                            """INSERT INTO THANH_TOAN (MaTT, MaLich, ThoiGianThanhToan, SoTien, PhuongThuc, TrangThaiTT)
                               VALUES (?, ?, ?, ?, ?, 'Đã thanh toán')""",
                            (matt, malich, now_str, sotien, phuongthuc),
                        )
                        sql_logs.append(
                            f"INSERT INTO THANH_TOAN (MaTT, MaLich, ThoiGianThanhToan, SoTien, PhuongThuc, TrangThaiTT) VALUES ('{matt}', '{malich}', '{now_str}', {int(sotien)}, '{phuongthuc}', 'Đã thanh toán');"
                        )
                        new_tt = {
                            "MaTT": matt,
                            "MaLich": malich,
                            "ThoiGianThanhToan": now_str,
                            "SoTien": sotien,
                            "PhuongThuc": phuongthuc,
                            "TrangThaiTT": "Đã thanh toán",
                        }
                else:
                    conn.close()
                    return self.send_json(
                        {"ok": False, "errors": [f"Vi phạm CHECK (CHK_TT_PhuongThuc): Phương thức '{phuongthuc}' không hợp lệ."]},
                        status=400,
                    )
                conn.commit()
            except sqlite3.Error as e:
                conn.close()
                return self.send_json({"ok": False, "errors": [f"Lỗi cập nhật THANH_TOAN: {str(e)}"]}, status=400)

            conn.close()
            return self.send_json(
                {
                    "ok": True,
                    "message": f"Đã cập nhật phương thức thanh toán cho đơn '{malich}' thành '{phuongthuc}'!",
                    "sql": "\n".join(sql_logs),
                    "thanh_toan": new_tt,
                }
            )


if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    port = 8080
    server = http.server.HTTPServer(("0.0.0.0", port), PetCareHandler)
    print(f"PetCare CSDL Server running at http://localhost:{port}")
    server.serve_forever()
