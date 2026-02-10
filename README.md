# ⚡ Auto Excel
## 🚀 Hướng dẫn cài đặt (Excel Desktop)

### Bước 1: Tải file Manifest
1. Nhấp chuột phải vào link này: **[manifest.xml](https://raw.githubusercontent.com/AliceAlicek2304/Excel-add-in/main/manifest.xml)**
2. Chọn **"Save link as..."** để tải về máy.

### Bước 2: Cài vào Excel
1. Bỏ file `manifest.xml` vào một thư mục (ví dụ: `C:\ExcelAddins`).
2. Chuột phải thư mục đó -> **Properties** -> **Sharing** -> **Share** -> Chọn **Everyone** -> **Add** -> **Share**.
3. Copy **Network Path** (Ví dụ: `\\TEN-MAY\ExcelAddins`).
4. Trong Excel: **File** -> **Options** -> **Trust Center** -> **Trust Center Settings** -> **Trusted Add-in Catalogs**.
5. Dán đường dẫn vào **Catalog URL** -> **Add catalog** -> Tích **Show in Menu** -> **OK**.
6. Khởi động lại Excel. Vào **Home** -> **Add-ins** -> **Shared Folder** để chọn Auto Excel.

---

## 🔑 Lấy Gemini API Key
1. Truy cập **[Google AI Studio](https://aistudio.google.com/app/apikey)** để lấy Key (Miễn phí).
2. Vào **Settings** trong Add-in và dán Key vào.

---

## 📖 Ví dụ câu lệnh:
- *"Tính tổng cột B điền vào C1"*
- *"Dò tên sản phẩm của mã ở ô D1 trong bảng A:B"*
- *"Lọc danh sách khách hàng có nợ > 0"*

---
*Repo: [AliceAlicek2304/Excel-add-in](https://github.com/AliceAlicek2304/Excel-add-in)*
