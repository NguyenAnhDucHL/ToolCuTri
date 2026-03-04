# 🗳️ Công cụ Tổng hợp Dữ liệu Cử tri

Ứng dụng web đơn giản chạy trên máy tính, giúp tự động tổng hợp dữ liệu cử tri từ nhiều file Excel vào bảng tổng hợp.

---

## 📋 Yêu cầu

- Python 3.9 trở lên
- pip (trình quản lý gói Python)

---

## 🚀 Cài đặt & Chạy

### Bước 1 — Cài thư viện

Mở Terminal (hoặc Command Prompt), chạy lệnh:

```bash
cd "/Users/macbookpro/Documents/tool dai bieu cu tri"
pip install -r requirements.txt
```

### Bước 2 — Khởi động ứng dụng

```bash
streamlit run app.py
```

Ứng dụng sẽ tự mở trình duyệt tại địa chỉ: **http://localhost:8501**

---

## 📖 Hướng dẫn sử dụng

### 1. Nhập đường dẫn folder nguồn

- **Đường dẫn local:** `/Users/ten/Documents/du_lieu_cu_tri`
- **Google Drive:** `https://drive.google.com/drive/folders/1Tzu3ymyE2NC...`

> ⚠️ Folder Google Drive phải được chia sẻ **"Anyone with the link"**

Mỗi file `.xlsx` trong folder = một khu phố/thôn/bản. **Tên file** (bỏ `.xlsx`) phải khớp với tên khu phố trong bảng tổng hợp.

### 2. Nhập đường dẫn file tổng hợp

File `BIỂU TỔNG HỢP DANH SÁCH CỬ TRI.xlsx` — có thể là đường dẫn local hoặc link Google Drive file.

### 3. Bấm "▶ Bắt đầu xử lý"

Hệ thống sẽ:
1. Đọc từng file `.xlsx` nguồn
2. Tìm hàng **Tổng/Nam/Nữ** tự động (hỗ trợ nhiều định dạng khác nhau)
3. Đếm cử tri 18 tuổi lần đầu (sinh 16/3/2007 – 15/3/2008)
4. Đếm cử tri cao tuổi >80 tuổi (sinh trước 15/3/1946)
5. Cập nhật vào bảng tổng hợp (cột F, G, H, K, L)

### 4. Tải file kết quả

Bấm nút **"⬇️ Tải file tổng hợp đã cập nhật"** để tải về máy.

---

## 📊 Cột được cập nhật

| Cột | Nội dung |
|-----|----------|
| F   | Tổng số cử tri |
| G   | Nam |
| H   | Nữ |
| K   | Cử tri 18 tuổi lần đầu tham gia bỏ phiếu |
| L   | Cử tri cao tuổi (trên 80 tuổi) |

---

## ⚠️ Lưu ý

- File gốc `BIỂU TỔNG HỢP` **không bị thay đổi** — bạn nhận bản mới sau khi xử lý.
- Nếu trùng tên file → chọn ngẫu nhiên 1 file.
- File nào lỗi sẽ được báo cáo, không ảnh hưởng các file khác.
- Cột trong file tổng hợp phải chứa tên khu phố khớp với tên file nguồn.

---

## 🏷️ Thông tin kỹ thuật

| Thư viện | Mục đích |
|----------|----------|
| Streamlit | Giao diện web |
| openpyxl | Đọc/ghi file Excel |
| pandas | Xử lý dữ liệu bảng |
| gdown | Tải file từ Google Drive |
| python-dateutil | Xử lý ngày tháng |
