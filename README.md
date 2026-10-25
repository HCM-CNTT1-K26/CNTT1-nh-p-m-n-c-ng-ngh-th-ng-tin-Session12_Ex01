# CNTT1-nh-p-m-n-c-ng-ngh-th-ng-tin-Session12_Ex01
---

### PHẦN 1. PHÂN TÍCH DATABASE

* **Record (Dòng):** Mỗi hàng là một giao dịch độc lập (ví dụ: giao dịch nạp tiền của sinh viên). Tuyệt đối không để dòng trống để tránh ngắt quãng dữ liệu.
* **Cột (Field):** Mỗi cột là một thuộc tính riêng (Mã GD, Số tiền, Tên...). Không gộp ô (Merge Cell) ở vùng dữ liệu để đảm bảo thao tác lọc, tra cứu chính xác.

### PHẦN 2. TRIỂN KHAI FILE

* **Word (SRS):** Cấu trúc 5 trang, dùng Auto TOC và phân cấp rõ Heading 1 (Tiêu đề lớn), Heading 2 (Tiêu đề nhỏ).
* **Excel:** Dùng công thức lọc chính xác kết hợp cắt khoảng trắng thừa:
`=VLOOKUP(TRIM(A2), 'DanhSach_User'!$A$2:$C$1000, 2, FALSE)`
* **PowerPoint:** Tiêu đề dạng khẳng định kết quả (Ví dụ: *"Doanh thu tăng 15% nhờ ưu đãi"*), dùng font chữ lớn (từ 30pt trở lên) và thay bảng số liệu bằng biểu đồ trực quan.

### PHẦN 3. XỬ LÝ LỖI #N/A TRONG VLOOKUP (4 BƯỚC)

1. **Khoảng trắng ẩn:** Dùng `TRIM()` để xóa dấu cách thừa ở mã giao dịch.
2. **Kiểu dữ liệu:** Kiểm tra xem dữ liệu bên tra cứu đang là dạng Chữ (Text) hay Số (Number).
3. **Vùng tham chiếu:** Cố định bảng tra cứu bằng phím `F4` (ví dụ: `$A$2:$C$100`) và đảm bảo giá trị cần tìm nằm ở cột đầu tiên của bảng.
4. **Tham số tìm kiếm:** Đảm bảo đối số cuối cùng của hàm VLOOKUP là `FALSE` (tìm chính xác tuyệt đối).

---

Bạn cần điều chỉnh thêm phần nào nữa không?
