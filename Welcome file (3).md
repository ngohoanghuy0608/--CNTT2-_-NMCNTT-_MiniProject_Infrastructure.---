## 1. Tạo cấu trúc thư mục

Mở PowerShell, chạy lệnh sau:

```
mkdir greentech-system
cd greentech-system
mkdir docs,scripts,data,backup
New-Item docs/README.md,docs/srs-he-thong.docx,data/log-he-thong.log,data/quan-ly-thiet-bi.xlsx,.env -ItemType File
```

Lưu ý: `.env` chỉ chứa mật khẩu giả định, không đưa mật khẩu thật lên GitHub.

## 2. Tạo script lọc Log

Tạo file `scripts/auto_report.bat` với nội dung:

```
@echo off
echo === BAO CAO SU CO GREENTECH ===
findstr /I "CRITICAL ERROR" data\log-he-thong.log >> backup\bao-cao-su-co.txt
echo Da loc log xong!
pause
```

Script này lọc các dòng chứa `CRITICAL` hoặc `ERROR` và ghi nối vào báo cáo.

Kiểm tra an toàn: Không chạy lệnh xóa dữ liệu nguy hiểm như `del /f /s /q` khi chưa hiểu rõ tác dụng.

## 3. Viết file README.md

Mở `docs/README.md`, viết nội dung mẫu:

```
# GreenTech System

## 1. Mục đích
Quản lý thiết bị IoT và theo dõi lỗi hệ thống.

## 2. Cách chạy script
```bat
scripts\auto_report.bat
```

## 3. Checklist kiểm tra
- [ ] Kiểm tra CPU
- [ ] Kiểm tra RAM
- [ ] Kiểm tra SSD
- [ ] Kiểm tra Log

## 4. Thông số phần cứng

| Thiết bị | Chức năng |
|---|---|
| CPU | Xử lý dữ liệu |
| RAM | Lưu dữ liệu tạm |
| SSD | Lưu trữ dữ liệu |

```

## 4. Tạo tài liệu Word

Mở Microsoft Word và tạo file `srs-he-thong.docx`.

-   Heading 1: Giới thiệu hệ thống.
    
-   Heading 2: Yêu cầu chức năng.
    
-   Heading 3: Chi tiết từng chức năng.
    
-   Chèn mục lục tự động bằng References → Table of Contents.
    

## 5. Làm bảng Excel

Tạo file `data/quan-ly-thiet-bi.xlsx` gồm 2 sheet:

-   `danh-muc-linh-kien`: Mã linh kiện, tên thiết bị, đơn giá, loại ổ cứng.
    
-   `trang-thai-tram`: Mã trạm, mã linh kiện, dung lượng RAM, ngày lắp đặt, tên thiết bị.
    

Công thức tra cứu tên thiết bị:

```
=VLOOKUP(TRIM(B2),'danh-muc-linh-kien'!$A$2:$D$100,2,FALSE)
```

Công thức giúp loại bỏ khoảng trắng thừa và tra cứu chính xác mã linh kiện.

## 6. Tính thời gian tải Log

Đổi tốc độ mạng từ Mbps sang MB/s bằng cách chia cho 8.

Ví dụ: mạng 80 Mbps tương đương 10 MB/s. File 500 MB tải trong khoảng 50 giây, chưa tính hao hụt mạng.
