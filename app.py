import http.server
import json
import sqlite3
import urllib.parse
import os
import io
import zipfile
from datetime import datetime, timedelta

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "petcare.db")
SQL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "csdl_cham_soc_thu_cung.sql")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def next_code(conn, table, col, prefix):
    rows = conn.execute(f"SELECT {col} FROM {table}").fetchall()
    max_n = 0
    for r in rows:
        digits = "".join(ch for ch in str(r[0]) if ch.isdigit())
        if digits:
            max_n = max(max_n, int(digits))
    return f"{prefix}{max_n + 1:03d}"


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
            return self.send_json({"results": rows})

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
            return self.send_json({"report1": rep1, "report2": rep2})

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

            conn = get_conn()
            if not makh:
                makh = next_code(conn, "KHACH_HANG", "MaKH", "KH")
            elif not makh.startswith("KH") or len(makh) < 4:
                errors.append("Mã khách hàng không hợp lệ.")
            if not hoten:
                errors.append("Vui lòng nhập họ và tên khách hàng.")
            if not sdt or not (10 <= len(sdt) <= 11) or not sdt.startswith("0") or not sdt.isdigit():
                errors.append("Số điện thoại không hợp lệ (vui lòng nhập 10–11 chữ số và bắt đầu bằng số 0).")

            if makh and conn.execute("SELECT 1 FROM KHACH_HANG WHERE MaKH = ?", (makh,)).fetchone():
                errors.append("Mã khách hàng này đã tồn tại trên hệ thống, vui lòng thử lại.")
            if sdt and conn.execute("SELECT 1 FROM KHACH_HANG WHERE SDT = ?", (sdt,)).fetchone():
                errors.append("Số điện thoại này đã được đăng ký cho một khách hàng khác.")
            if email and conn.execute("SELECT 1 FROM KHACH_HANG WHERE LOWER(Email) = LOWER(?)", (email,)).fetchone():
                errors.append("Địa chỉ email này đã được đăng ký cho một khách hàng khác.")

            if errors:
                conn.close()
                return self.send_json({"ok": False, "errors": errors}, status=400)

            try:
                conn.execute(
                    "INSERT INTO KHACH_HANG (MaKH, HoTen, SDT, Email, DiaChi) VALUES (?, ?, ?, ?, ?)",
                    (makh, hoten, sdt, email, diachi),
                )
                conn.commit()
            except sqlite3.IntegrityError:
                conn.close()
                return self.send_json(
                    {"ok": False, "errors": ["Không thể lưu thông tin do dữ liệu chưa hợp lệ hoặc bị trùng lặp."]},
                    status=400,
                )
            conn.close()
            return self.send_json(
                {
                    "ok": True,
                    "message": f"Đã lưu thông tin khách hàng '{hoten}' thành công!",
                    "MaKH": makh,
                }
            )

        elif parsed.path == "/api/thu_cung":
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
            kh_hoten = kh_sdt = kh_email = kh_diachi = ""

            if new_kh and isinstance(new_kh, dict):
                makh = str(new_kh.get("MaKH", "")).strip()
                kh_hoten = str(new_kh.get("HoTen", "")).strip()
                kh_sdt = str(new_kh.get("SDT", "")).strip()
                kh_email = str(new_kh.get("Email", "")).strip() or None
                kh_diachi = str(new_kh.get("DiaChi", "")).strip()

                if not makh:
                    makh = next_code(conn, "KHACH_HANG", "MaKH", "KH")
                elif not makh.startswith("KH") or len(makh) < 4:
                    errors.append("Mã khách hàng không hợp lệ.")
                if not kh_hoten:
                    errors.append("Vui lòng nhập họ và tên khách hàng.")
                if not kh_sdt or not (10 <= len(kh_sdt) <= 11) or not kh_sdt.startswith("0") or not kh_sdt.isdigit():
                    errors.append("Số điện thoại không hợp lệ (vui lòng nhập 10–11 chữ số và bắt đầu bằng số 0).")
                if makh and conn.execute("SELECT 1 FROM KHACH_HANG WHERE MaKH = ?", (makh,)).fetchone():
                    errors.append("Mã khách hàng này đã tồn tại trên hệ thống, vui lòng thử lại.")
                if kh_sdt and conn.execute("SELECT 1 FROM KHACH_HANG WHERE SDT = ?", (kh_sdt,)).fetchone():
                    errors.append("Số điện thoại này đã được đăng ký cho một khách hàng khác.")
                if kh_email and conn.execute("SELECT 1 FROM KHACH_HANG WHERE LOWER(Email) = LOWER(?)", (kh_email,)).fetchone():
                    errors.append("Địa chỉ email này đã được đăng ký cho một khách hàng khác.")

            if not matc:
                matc = next_code(conn, "THU_CUNG", "MaThuCung", "TC")
            elif not matc.startswith("TC") or len(matc) < 4:
                errors.append("Mã thú cưng không hợp lệ.")
            if not makh:
                errors.append("Vui lòng chọn hoặc nhập thông tin chủ nuôi.")
            if not tentc:
                errors.append("Vui lòng nhập tên thú cưng.")
            if not loai:
                errors.append("Vui lòng chọn loài thú cưng.")
            if gioitinh not in ("Đực", "Cái", "Chưa xác định"):
                errors.append("Giới tính thú cưng chỉ nhận: Đực, Cái hoặc Chưa xác định.")

            try:
                tuoi = int(body.get("Tuoi", 0))
                if tuoi <= 0 or tuoi > 360:
                    errors.append("Độ tuổi thú cưng cần nằm trong khoảng từ 1 đến 360 tháng.")
            except (ValueError, TypeError):
                errors.append("Độ tuổi thú cưng cần nằm trong khoảng từ 1 đến 360 tháng.")
                tuoi = 0

            try:
                cannang = float(body.get("CanNang", 0))
                if cannang <= 0 or cannang > 150:
                    errors.append("Cân nặng thú cưng cần nằm trong khoảng từ 0.1 đến 150 kg.")
            except (ValueError, TypeError):
                errors.append("Cân nặng thú cưng cần nằm trong khoảng từ 0.1 đến 150 kg.")
                cannang = 0

            if matc:
                dup = conn.execute("SELECT 1 FROM THU_CUNG WHERE MaThuCung = ?", (matc,)).fetchone()
                if dup:
                    errors.append("Hồ sơ thú cưng này đã tồn tại trên hệ thống, vui lòng thử lại.")
            if makh and not new_kh:
                kh_ok = conn.execute("SELECT 1 FROM KHACH_HANG WHERE MaKH = ?", (makh,)).fetchone()
                if not kh_ok:
                    errors.append("Không tìm thấy hồ sơ khách hàng tương ứng.")

            if errors:
                conn.close()
                return self.send_json({"ok": False, "errors": errors}, status=400)

            try:
                if new_kh and isinstance(new_kh, dict):
                    conn.execute(
                        "INSERT INTO KHACH_HANG (MaKH, HoTen, SDT, Email, DiaChi) VALUES (?, ?, ?, ?, ?)",
                        (makh, kh_hoten, kh_sdt, kh_email, kh_diachi),
                    )
                sql = """INSERT INTO THU_CUNG (MaThuCung, MaKH, TenThuCung, Loai, Giong, GioiTinh, Tuoi, CanNang, GhiChu)
                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"""
                conn.execute(sql, (matc, makh, tentc, loai, giong, gioitinh, tuoi, cannang, ghichu))
                conn.commit()
            except sqlite3.IntegrityError:
                conn.close()
                return self.send_json(
                    {"ok": False, "errors": ["Không thể lưu hồ sơ thú cưng do dữ liệu chưa hợp lệ hoặc bị trùng lặp."]},
                    status=400,
                )
            conn.close()
            return self.send_json(
                {
                    "ok": True,
                    "message": f"Đã lưu hồ sơ thú cưng '{tentc}' thành công!",
                    "MaThuCung": matc,
                    "MaKH": makh,
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
            combo_madv = str(body.get("ComboMaDV", "")).strip()
            combo_malich = str(body.get("ComboMaLich", "")).strip()
            try:
                combo_soluong = int(body.get("ComboSoLuong", 1))
            except (ValueError, TypeError):
                combo_soluong = 1

            conn = get_conn()
            if not malich:
                malich = next_code(conn, "LICH_HEN", "MaLich", "LH")
            elif not malich.startswith("LH") or len(malich) < 4:
                errors.append("Mã lịch hẹn không hợp lệ.")
            if not matc:
                errors.append("Vui lòng chọn thú cưng cần chăm sóc.")
            if not madv:
                errors.append("Vui lòng chọn gói dịch vụ.")
            if not manv:
                errors.append("Vui lòng chọn chuyên viên phụ trách.")
            if not ngaygio:
                errors.append("Vui lòng chọn ngày và giờ hẹn.")

            try:
                soluong = int(body.get("SoLuong", 0))
                if soluong <= 0:
                    errors.append("Số lượng dịch vụ phải lớn hơn 0.")
            except (ValueError, TypeError):
                errors.append("Số lượng dịch vụ phải lớn hơn 0.")
                soluong = 0

            if malich:
                if conn.execute("SELECT 1 FROM LICH_HEN WHERE MaLich = ?", (malich,)).fetchone():
                    errors.append("Mã lịch hẹn này đã tồn tại, vui lòng thử lại.")
            if matc:
                if not conn.execute("SELECT 1 FROM THU_CUNG WHERE MaThuCung = ?", (matc,)).fetchone():
                    errors.append("Không tìm thấy hồ sơ thú cưng đã chọn.")

            dv_row = None
            if madv:
                dv_row = conn.execute("SELECT MaDV, TenDV, DonGia, ThoiLuongDuKien FROM DICH_VU WHERE MaDV = ?", (madv,)).fetchone()
                if not dv_row:
                    errors.append("Không tìm thấy gói dịch vụ đã chọn.")

            combo_dv_row = None
            if combo_madv:
                if combo_madv == madv:
                    errors.append("Gói dịch vụ đi kèm không được trùng với dịch vụ chính (hãy tăng số lượng nếu dùng cùng một gói).")
                combo_dv_row = conn.execute("SELECT MaDV, TenDV, DonGia, ThoiLuongDuKien FROM DICH_VU WHERE MaDV = ?", (combo_madv,)).fetchone()
                if not combo_dv_row:
                    errors.append("Không tìm thấy gói dịch vụ đi kèm đã chọn.")
                if combo_soluong <= 0:
                    errors.append("Số lượng dịch vụ đi kèm phải lớn hơn 0.")

            def calc_staff_minutes(dv_code, thoi_luong, sl):
                if dv_code == "DV005":
                    return 60
                return int(thoi_luong or 60) * max(1, int(sl or 1))

            dur1 = calc_staff_minutes(madv, dv_row["ThoiLuongDuKien"] if dv_row else 60, soluong)
            dur2 = calc_staff_minutes(combo_madv, combo_dv_row["ThoiLuongDuKien"] if combo_dv_row else 0, combo_soluong) if combo_dv_row else 0
            total_dur = dur1 + dur2

            if manv:
                nv_row = conn.execute("SELECT HoTen, TrangThai, CaLamViec, ChucVu FROM NHAN_VIEN WHERE MaNV = ?", (manv,)).fetchone()
                if not nv_row:
                    errors.append("Không tìm thấy chuyên viên đã chọn.")
                else:
                    if nv_row["TrangThai"] != "Đang làm việc":
                        errors.append(
                            f"Chuyên viên {nv_row['HoTen']} hiện đang {nv_row['TrangThai'].lower()}, vui lòng chọn chuyên viên khác."
                        )
                    dt_start = None
                    if ngaygio and len(ngaygio) >= 16:
                        try:
                            dt_start = datetime.strptime(ngaygio[:16], "%Y-%m-%d %H:%M")
                            dt_end = dt_start + timedelta(minutes=total_dur)
                            hr_start = dt_start.hour + dt_start.minute / 60.0
                            hr_end = dt_end.hour + dt_end.minute / 60.0 + (24.0 if dt_end.date() > dt_start.date() else 0.0)
                            ca = nv_row["CaLamViec"]
                            s_min, s_max = 8.0, 21.0
                            if "Ca sáng" in ca:
                                s_min, s_max = 8.0, 12.0
                            elif "Ca chiều" in ca:
                                s_min, s_max = 13.0, 17.0
                            elif "Ca tối" in ca:
                                s_min, s_max = 17.0, 21.0
                            elif "Hành chính" in ca:
                                s_min, s_max = 8.0, 17.0

                            if not (s_min <= hr_start < s_max):
                                errors.append(
                                    f"Chuyên viên {nv_row['HoTen']} làm việc {ca}, chưa phù hợp với giờ hẹn {ngaygio[11:16]}. Vui lòng chọn giờ khác hoặc chuyên viên khác."
                                )
                            elif hr_end > s_max:
                                errors.append(
                                    f"Tổng thời gian chăm sóc dự kiến là {total_dur} phút ({dt_start.strftime('%H:%M')} – {dt_end.strftime('%H:%M')}), vượt quá giờ kết thúc ca làm việc ({ca})."
                                )
                        except ValueError:
                            errors.append("Định dạng ngày giờ hẹn chưa hợp lệ.")

                    if dt_start is not None:
                        dt_end = dt_start + timedelta(minutes=total_dur)
                        existing_bookings = conn.execute(
                            """SELECT lh.MaLich, lh.NgayGioHen, lh.MaDV, lh.SoLuong, dv.TenDV, dv.ThoiLuongDuKien
                               FROM LICH_HEN lh
                               JOIN DICH_VU dv ON lh.MaDV = dv.MaDV
                               WHERE lh.MaNV = ? AND lh.TrangThai NOT IN ('Đã hủy', 'Không đến')""",
                            (manv,),
                        ).fetchall()
                        for ex in existing_bookings:
                            try:
                                ex_start = datetime.strptime(ex["NgayGioHen"][:16], "%Y-%m-%d %H:%M")
                                ex_dur = calc_staff_minutes(ex["MaDV"], ex["ThoiLuongDuKien"], ex["SoLuong"])
                                ex_end = ex_start + timedelta(minutes=ex_dur)
                                if dt_start < ex_end and ex_start < dt_end:
                                    errors.append(
                                        f"Chuyên viên {nv_row['HoTen']} đã kín lịch trong khung giờ {ex_start.strftime('%H:%M')} – {ex_end.strftime('%H:%M')} ngày {ex_start.strftime('%d/%m/%Y')}. Vui lòng chọn khung giờ khác hoặc chuyên viên khác."
                                    )
                                    break
                            except ValueError:
                                pass

            if errors:
                conn.close()
                return self.send_json({"ok": False, "errors": errors}, status=400)

            created_bookings = []
            try:
                conn.execute(
                    """INSERT INTO LICH_HEN (MaLich, MaThuCung, MaDV, MaNV, NgayGioHen, SoLuong, TrangThai, GhiChu)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                    (malich, matc, madv, manv, ngaygio, soluong, trangthai, ghichu),
                )
                created_bookings.append(malich)

                if combo_madv and combo_dv_row:
                    if not combo_malich or combo_malich == malich:
                        combo_malich = next_code(conn, "LICH_HEN", "MaLich", "LH")
                    dt_combo = datetime.strptime(ngaygio[:16], "%Y-%m-%d %H:%M") + timedelta(minutes=dur1)
                    ngaygio_combo = dt_combo.strftime("%Y-%m-%d %H:%M:00")
                    ghichu_combo = f"[Combo sau #{malich}] {ghichu}".strip()
                    conn.execute(
                        """INSERT INTO LICH_HEN (MaLich, MaThuCung, MaDV, MaNV, NgayGioHen, SoLuong, TrangThai, GhiChu)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                        (combo_malich, matc, combo_madv, manv, ngaygio_combo, combo_soluong, trangthai, ghichu_combo),
                    )
                    created_bookings.append(combo_malich)

                conn.commit()
            except sqlite3.IntegrityError:
                conn.close()
                return self.send_json(
                    {"ok": False, "errors": ["Không thể đặt lịch hẹn do dữ liệu chưa hợp lệ."]},
                    status=400,
                )

            conn.close()
            return self.send_json(
                {
                    "ok": True,
                    "message": f"Đặt lịch hẹn thành công! Mã đặt lịch: {', '.join('#' + x for x in created_bookings)}.",
                    "created_bookings": created_bookings,
                }
            )

        elif parsed.path == "/api/lich_hen/status":
            malich = str(body.get("MaLich", "")).strip()
            trangthai = str(body.get("TrangThai", "")).strip()
            phuongthuc = str(body.get("PhuongThuc", "Chuyển khoản")).strip()

            valid_statuses = ("Mới tạo", "Đã xác nhận", "Đang thực hiện", "Hoàn thành", "Đã hủy", "Không đến")
            if not malich:
                return self.send_json({"ok": False, "errors": ["Không tìm thấy mã lịch hẹn."]}, status=400)
            if trangthai not in valid_statuses:
                return self.send_json(
                    {"ok": False, "errors": ["Trạng thái lịch hẹn không hợp lệ."]},
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
                return self.send_json({"ok": False, "errors": [f"Không tìm thấy lịch hẹn #{malich}."]}, status=404)

            try:
                conn.execute("UPDATE LICH_HEN SET TrangThai = ? WHERE MaLich = ?", (trangthai, malich))

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
                    else:
                        matt = next_code(conn, "THANH_TOAN", "MaTT", "TT")
                        conn.execute(
                            """INSERT INTO THANH_TOAN (MaTT, MaLich, ThoiGianThanhToan, SoTien, PhuongThuc, TrangThaiTT)
                               VALUES (?, ?, ?, ?, ?, 'Đã thanh toán')""",
                            (matt, malich, now_str, sotien, phuongthuc),
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

                conn.commit()
            except sqlite3.Error:
                conn.close()
                return self.send_json({"ok": False, "errors": ["Không thể cập nhật trạng thái lịch hẹn."]}, status=400)

            conn.close()
            return self.send_json(
                {
                    "ok": True,
                    "message": f"Đã cập nhật trạng thái lịch hẹn #{malich} sang '{trangthai}'.",
                    "thanh_toan": new_tt,
                }
            )

        elif parsed.path == "/api/thanh_toan/method":
            malich = str(body.get("MaLich", "")).strip()
            phuongthuc = str(body.get("PhuongThuc", "")).strip()

            if not malich:
                return self.send_json({"ok": False, "errors": ["Không tìm thấy mã lịch hẹn."]}, status=400)

            conn = get_conn()
            lh_row = conn.execute(
                """SELECT lh.MaLich, lh.SoLuong, lh.TrangThai, dv.DonGia
                   FROM LICH_HEN lh JOIN DICH_VU dv ON lh.MaDV = dv.MaDV
                   WHERE lh.MaLich = ?""",
                (malich,),
            ).fetchone()
            if not lh_row:
                conn.close()
                return self.send_json({"ok": False, "errors": [f"Không tìm thấy lịch hẹn #{malich}."]}, status=404)

            if lh_row["TrangThai"] in ("Đã hủy", "Không đến") and phuongthuc != "Chưa thanh toán":
                conn.close()
                return self.send_json(
                    {
                        "ok": False,
                        "errors": [
                            "Lịch hẹn đã hủy hoặc khách không đến nên không thể ghi nhận thu tiền."
                        ],
                    },
                    status=400,
                )

            new_tt = None
            updated_lh_status = None
            try:
                existing_tt = conn.execute("SELECT MaTT FROM THANH_TOAN WHERE MaLich = ?", (malich,)).fetchone()
                if phuongthuc == "Chưa thanh toán":
                    if existing_tt:
                        conn.execute("UPDATE THANH_TOAN SET TrangThaiTT = 'Chưa thanh toán' WHERE MaLich = ?", (malich,))
                elif phuongthuc in ("Tiền mặt", "Chuyển khoản", "Thẻ"):
                    sotien = float(lh_row["DonGia"]) * int(lh_row["SoLuong"])
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    if lh_row["TrangThai"] != "Hoàn thành":
                        conn.execute("UPDATE LICH_HEN SET TrangThai = 'Hoàn thành' WHERE MaLich = ?", (malich,))
                        updated_lh_status = "Hoàn thành"
                    if existing_tt:
                        conn.execute(
                            "UPDATE THANH_TOAN SET PhuongThuc = ?, TrangThaiTT = 'Đã thanh toán', SoTien = ?, ThoiGianThanhToan = ? WHERE MaLich = ?",
                            (phuongthuc, sotien, now_str, malich),
                        )
                    else:
                        matt = next_code(conn, "THANH_TOAN", "MaTT", "TT")
                        conn.execute(
                            """INSERT INTO THANH_TOAN (MaTT, MaLich, ThoiGianThanhToan, SoTien, PhuongThuc, TrangThaiTT)
                               VALUES (?, ?, ?, ?, ?, 'Đã thanh toán')""",
                            (matt, malich, now_str, sotien, phuongthuc),
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
                        {"ok": False, "errors": ["Hình thức thanh toán không hợp lệ."]},
                        status=400,
                    )
                conn.commit()
            except sqlite3.Error:
                conn.close()
                return self.send_json({"ok": False, "errors": ["Không thể cập nhật thông tin thanh toán."]}, status=400)

            conn.close()
            return self.send_json(
                {
                    "ok": True,
                    "message": f"Đã cập nhật thanh toán cho lịch hẹn #{malich} ({phuongthuc}).",
                    "thanh_toan": new_tt,
                    "updated_lh_status": updated_lh_status,
                }
            )

        elif parsed.path == "/api/khach_hang/update":
            errors = []
            makh = str(body.get("MaKH", "")).strip()
            hoten = str(body.get("HoTen", "")).strip()
            sdt = str(body.get("SDT", "")).strip()
            email = str(body.get("Email", "")).strip() or None
            diachi = str(body.get("DiaChi", "")).strip()

            if not makh:
                errors.append("Không xác định được khách hàng cần cập nhật.")
            if not hoten:
                errors.append("Vui lòng nhập họ và tên khách hàng.")
            if not sdt or not (10 <= len(sdt) <= 11) or not sdt.startswith("0") or not sdt.isdigit():
                errors.append("Số điện thoại không hợp lệ (vui lòng nhập 10–11 chữ số và bắt đầu bằng số 0).")

            conn = get_conn()
            if makh and not conn.execute("SELECT 1 FROM KHACH_HANG WHERE MaKH = ?", (makh,)).fetchone():
                errors.append("Không tìm thấy hồ sơ khách hàng.")
            if sdt and conn.execute("SELECT 1 FROM KHACH_HANG WHERE SDT = ? AND MaKH != ?", (sdt, makh)).fetchone():
                errors.append("Số điện thoại này đã được đăng ký cho một khách hàng khác.")
            if email and conn.execute("SELECT 1 FROM KHACH_HANG WHERE LOWER(Email) = LOWER(?) AND MaKH != ?", (email, makh)).fetchone():
                errors.append("Địa chỉ email này đã được đăng ký cho một khách hàng khác.")

            if errors:
                conn.close()
                return self.send_json({"ok": False, "errors": errors}, status=400)

            try:
                conn.execute(
                    "UPDATE KHACH_HANG SET HoTen = ?, SDT = ?, Email = ?, DiaChi = ? WHERE MaKH = ?",
                    (hoten, sdt, email, diachi, makh),
                )
                conn.commit()
            except sqlite3.Error:
                conn.close()
                return self.send_json({"ok": False, "errors": ["Không thể cập nhật thông tin khách hàng."]}, status=400)
            conn.close()
            return self.send_json(
                {
                    "ok": True,
                    "message": f"Đã cập nhật thông tin khách hàng '{hoten}' thành công!",
                }
            )

        elif parsed.path == "/api/thu_cung/update":
            errors = []
            matc = str(body.get("MaThuCung", "")).strip()
            makh = str(body.get("MaKH", "")).strip()
            tentc = str(body.get("TenThuCung", "")).strip()
            loai = str(body.get("Loai", "")).strip()
            giong = str(body.get("Giong", "")).strip()
            gioitinh = str(body.get("GioiTinh", "")).strip()
            ghichu = str(body.get("GhiChu", "")).strip()

            if not matc:
                errors.append("Không xác định được hồ sơ thú cưng cần cập nhật.")
            if not makh:
                errors.append("Vui lòng chọn chủ nuôi cho thú cưng.")
            if not tentc:
                errors.append("Vui lòng nhập tên thú cưng.")
            if not loai:
                errors.append("Vui lòng chọn loài thú cưng.")
            if gioitinh not in ("Đực", "Cái", "Chưa xác định"):
                errors.append("Giới tính thú cưng chỉ nhận: Đực, Cái hoặc Chưa xác định.")

            try:
                tuoi = int(body.get("Tuoi", 0))
                if tuoi <= 0 or tuoi > 360:
                    errors.append("Độ tuổi thú cưng cần nằm trong khoảng từ 1 đến 360 tháng.")
            except (ValueError, TypeError):
                errors.append("Độ tuổi thú cưng cần nằm trong khoảng từ 1 đến 360 tháng.")
                tuoi = 0

            try:
                cannang = float(body.get("CanNang", 0))
                if cannang <= 0 or cannang > 150:
                    errors.append("Cân nặng thú cưng cần nằm trong khoảng từ 0.1 đến 150 kg.")
            except (ValueError, TypeError):
                errors.append("Cân nặng thú cưng cần nằm trong khoảng từ 0.1 đến 150 kg.")
                cannang = 0

            conn = get_conn()
            if matc and not conn.execute("SELECT 1 FROM THU_CUNG WHERE MaThuCung = ?", (matc,)).fetchone():
                errors.append("Không tìm thấy hồ sơ thú cưng.")
            if makh and not conn.execute("SELECT 1 FROM KHACH_HANG WHERE MaKH = ?", (makh,)).fetchone():
                errors.append("Không tìm thấy hồ sơ khách hàng tương ứng.")

            if errors:
                conn.close()
                return self.send_json({"ok": False, "errors": errors}, status=400)

            try:
                conn.execute(
                    """UPDATE THU_CUNG
                       SET MaKH = ?, TenThuCung = ?, Loai = ?, Giong = ?, GioiTinh = ?, Tuoi = ?, CanNang = ?, GhiChu = ?
                       WHERE MaThuCung = ?""",
                    (makh, tentc, loai, giong, gioitinh, tuoi, cannang, ghichu, matc),
                )
                conn.commit()
            except sqlite3.Error:
                conn.close()
                return self.send_json({"ok": False, "errors": ["Không thể cập nhật hồ sơ thú cưng."]}, status=400)
            conn.close()
            return self.send_json(
                {
                    "ok": True,
                    "message": f"Đã cập nhật hồ sơ thú cưng '{tentc}' thành công!",
                }
            )

        elif parsed.path == "/api/nhan_vien/update":
            manv = str(body.get("MaNV", "")).strip()
            trangthai = str(body.get("TrangThai", "")).strip()
            calamviec = str(body.get("CaLamViec", "")).strip()
            valid_st = ("Đang làm việc", "Nghỉ phép", "Đã nghỉ việc")
            valid_ca = ("Ca sáng (08:00-12:00)", "Ca chiều (13:00-17:00)", "Ca tối (17:00-21:00)", "Hành chính (08:00-17:00)")
            if trangthai not in valid_st:
                return self.send_json({"ok": False, "errors": ["Trạng thái làm việc của nhân viên không hợp lệ."]}, status=400)
            if calamviec not in valid_ca:
                return self.send_json({"ok": False, "errors": ["Ca làm việc không hợp lệ."]}, status=400)
            conn = get_conn()
            conn.execute("UPDATE NHAN_VIEN SET TrangThai = ?, CaLamViec = ? WHERE MaNV = ?", (trangthai, calamviec, manv))
            conn.commit()
            conn.close()
            return self.send_json(
                {
                    "ok": True,
                    "message": "Đã cập nhật thông tin nhân viên thành công!",
                }
            )

        elif parsed.path == "/api/dich_vu/update":
            madv = str(body.get("MaDV", "")).strip()
            try:
                dongia = float(body.get("DonGia", -1))
                thoiluong = int(body.get("ThoiLuongDuKien", 0))
            except (ValueError, TypeError):
                return self.send_json({"ok": False, "errors": ["Đơn giá và thời lượng dịch vụ phải là số hợp lệ."]}, status=400)
            if dongia < 0 or thoiluong <= 0:
                return self.send_json({"ok": False, "errors": ["Đơn giá không được âm và thời lượng dịch vụ phải lớn hơn 0 phút."]}, status=400)
            conn = get_conn()
            conn.execute("UPDATE DICH_VU SET DonGia = ?, ThoiLuongDuKien = ? WHERE MaDV = ?", (dongia, thoiluong, madv))
            conn.commit()
            conn.close()
            return self.send_json(
                {
                    "ok": True,
                    "message": "Đã cập nhật thông tin dịch vụ thành công!",
                }
            )


if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    port = 8080
    server = http.server.HTTPServer(("0.0.0.0", port), PetCareHandler)
    print(f"PetCare Spa & Clinic Server running at http://localhost:{port}")
    server.serve_forever()
