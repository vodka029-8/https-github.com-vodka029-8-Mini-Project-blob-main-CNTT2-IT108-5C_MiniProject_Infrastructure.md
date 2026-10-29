## 1. Cấu trúc Cây thư mục dự án (Kebab-case)

Bạn hãy mở Terminal hoặc Git Bash tại thư mục muốn làm việc và chạy lệnh sau để khởi tạo nhanh cấu trúc:

```bash
mkdir -p greentech-system/docs greentech-system/scripts greentech-system/data greentech-system/backup

```

Cấu trúc thư mục chuẩn sau khi tạo sẽ như sau:

```text
[TênLớp]_[MônHọc]_MiniProject_Infrastructure/
├── docs/
│   ├── README.md
│   └── srs-he-thong.docx
├── scripts/
│   └── auto_report.bat
├── data/
│   ├── log-he-thong.log
│   └── quan-ly-thiet-bi.xlsx
└── .env

```

---

## 2. Module 1: Thiết lập "Căn cứ số" (Terminal)

* **Khởi tạo dự án:** Sử dụng các lệnh `cd` để di chuyển vào thư mục dự án.
* **Tạo file cấu hình ẩn `.env`:**
* Trên Windows (CMD/PowerShell): Tạo file và nhập nội dung cấu hình giả định như mật khẩu cơ sở dữ liệu, API key.
* Ví dụ nội dung file `.env`:
```env
DB_HOST=localhost
DB_PORT=5432
SYSTEM_USER=admin_greentech
SECRET_KEY=greentech_secure_2026

```





---

## 3. Module 2: Robot lọc Log tự động (`auto_report.bat`)

Bạn hãy tạo file `scripts/auto_report.bat` với nội dung script tự động hóa lọc lỗi từ file log mà không làm tràn RAM:

```batch
@echo off
chcp 65001 >nul
echo ==================================================== >> ..\backup\bao-cao-su-co.txt
echo BAO CAO SU CO HE THONG - %DATE% %TIME% >> ..\backup\bao-cao-su-co.txt
echo ==================================================== >> ..\backup\bao-cao-su-co.txt

REM Su dung pipe va findstr de loc cac dong CRITICAL hoac ERROR tu file log lon
type ..\data\log-he-thong.log | findstr /I "CRITICAL ERROR" >> ..\backup\bao-cao-su-co.txt

echo Da hoang tat trich xuat log su co vao thu muc backup!
pause

```

---

## 4. Module 3: Tài liệu hướng dẫn vận hành (`README.md` & Word)

### Nội dung mẫu cho `docs/README.md`:

```markdown
# Hệ Thống Quản Trị Hạ Tầng & Nhật Ký Kỹ Thuật - GreenTech

## 1. Tổng quan hệ thống
Hệ thống quản lý hạ tầng cho Startup công nghệ **GreenTech**, giám sát log tự động cho các trạm cảm biến IoT ngoài trang trại.

## 2. Hướng dẫn chạy Script tự động hóa
Truyền lệnh thực thi script trích xuất log sự cố:
```bash
cd scripts
auto_report.bat

```

## 3. Checklist kiểm tra thiết bị phần cứng

* [x] Khởi tạo cây thư mục chuẩn kebab-case
* [ ] Kiểm tra dung lượng RAM 100 trạm cảm biến
* [ ] Đối soát dữ liệu asset phần cứng qua Excel
* [ ] Backup báo cáo lỗi hệ thống định kỳ

## 4. Bảng thông số phần cứng tham chiếu

| Mã Linh Kiện | Tên Thiết Bị | Đơn Giá (VNĐ) | Loại Storage |
| --- | --- | --- | --- |
| CPU-I7-10 | Intel Core i7-10850H | 4,500,000 | N/A |
| SSD-512-NV | SSD NVMe 512GB | 1,200,000 | SSD |
| RAM-16-D4 | RAM 16GB DDR4 | 950,000 | N/A |

```

*   **Đối với tài liệu SRS (`srs-he-thong.docx`):** Bạn hãy mở Microsoft Word, áp dụng chuẩn Heading Styles (Heading 1 cho tiêu đề lớn, Heading 2 cho phần nhỏ), sau đó chèn bảng mục lục tự động (**Auto TOC**) ở trang đầu tiên.

---

## 5. Module 4: Đối soát tài nguyên trên Excel (`quan-ly-thiet-bi.xlsx`)
Bạn tạo file Excel đặt trong thư mục `data/` gồm 2 Sheet:
1.  **Sheet `danh-muc-linh-kien`**: Gồm các cột `ma_linh_kien` (Khóa chính), `ten_thiet_bi`, `don_gia`, `loai_storage`.
2.  **Sheet `trang-thai-tram`**: Gồm các cột `ma_tram`, `ma_linh_kien`, `dung_luong_ram`, `ngay_lap_dat`.
    *   Tại cột điền tên thiết bị trong sheet `trang-thai-tram`, sử dụng hàm **VLOOKUP** kết hợp cố định vùng dữ liệu bằng dấu `$` và tham số `FALSE` để đối soát chính xác:
        ```excel
        =VLOOKUP(B2, danh-muc-linh-kien!$A$2:$D$10, 2, FALSE)
