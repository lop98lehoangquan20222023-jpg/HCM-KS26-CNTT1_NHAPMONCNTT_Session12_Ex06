# BÀI LÀM: BÁO CÁO SỰ CỐ AN NINH MẠNG NETGUARD
## PHẦN 1. PHÂN TÍCH

### Cơ chế "Google Search của bảng tính" thông qua AutoFilter trong Excel

Theo em hiểu, tính năng AutoFilter trong Excel giống như một công cụ tìm kiếm Google thu nhỏ vì nó có những cách hoạt động tương tự như sau:

1. Lọc từ **khóa giống như gõ ô tìm kiếm:**

   * Khi bật AutoFilter, ở đầu mỗi cột sẽ xuất hiện một nút bấm. Bấm vào đó sẽ có một ô `Search`. Khi mình gõ chữ gì vào (ví dụ: `Ransomware` hay `Trojan`), Excel sẽ tự lọc ra đúng các dòng có chứa từ đó, giống hệt việc gõ từ khóa lên Google để tìm bài viết.

2. **Tìm kiếm theo điều kiện (Text Filters):**

   * Trong Google mình có thể tìm kiếm nâng cao thì trong Excel cũng có mục `Text Filters`. Mình có thể chọn các điều kiện như `Begins With...` (bắt đầu bằng chữ gì đó) hoặc `Contains...` (chứa chữ gì đó) để tìm đúng dữ liệu mình cần.

3. **Gom nhóm dữ liệu (Checklist):**

   * AutoFilter tự động quét từ trên xuống dưới cột đó rồi gom tất cả các giá trị trùng nhau lại thành một danh sách. Mình chỉ cần tích chọn hoặc bỏ tích là có thể xem đúng phần dữ liệu mình muốn, không phải lội qua hàng nghìn dòng dữ liệu thô.

## PHẦN 2. TRIỂN KHAI

### 1. Xử lý tra cứu mã độc trong Excel bằng VLOOKUP kết hợp TRIM

Trong thực tế, khi xuất file log (nhật ký hệ thống) ra Excel, các mã độc hay bị dính khoảng trắng ở đầu hoặc ở cuối do lỗi nhập liệu (ví dụ: `"MAL_001 "` thay vì `"MAL_001"`). Nếu dùng hàm `VLOOKUP` bình thường sẽ bị lỗi `#N/A` vì Excel tưởng đó là 2 chuỗi khác nhau.

**Cách xử lý:** Em sẽ lồng thêm hàm `TRIM` vào vị trí giá trị tìm kiếm (`Lookup_value`) để cắt bỏ hết khoảng trắng thừa trước khi tra cứu.

* **Công thức tại ô C2 (Cột Tên mã độc):**

  ```
  =VLOOKUP(TRIM(B2), 'Cơ sở dữ liệu virus'!$A$2:$C$50, 2, FALSE)
  
  ```

* **Giải thích các thành phần trong công thức:**

  * `TRIM(B2)`: Hàm này sẽ tự động xóa các khoảng trắng thừa ở đầu và đuôi của ô B2 (ô chứa mã độc bị lỗi khoảng trắng).

  * `'Cơ sở dữ liệu virus'!$A$2:$C$50`: Bảng danh mục mã độc chuẩn dùng để đối chiếu. Cần ấn `F4` để cố định vùng này lại (\$) khi kéo công thức xuống không bị lệch.

  * `2`: Lấy dữ liệu ở cột thứ 2 (Tên loại mã độc) trong bảng CSDL.

  * `FALSE`: Tìm kiếm chính xác 100%.

### 2. Định dạng báo cáo Word & Tạo mục lục tự động (Auto TOC)

Để làm một bài báo cáo sự cố dài