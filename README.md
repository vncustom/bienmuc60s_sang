# App biên mục 60sa

Ứng dụng tạo file Map tổng cho bản tin 60s sáng hoặc 60s tối từ thư mục input gồm playlist Excel và các file kịch bản RTF.

## Cài đặt

```powershell
pip install -r requirements.txt
```

## Chạy app

```powershell
python app_bien_muc_60s.py
```

Trên giao diện:

- Chọn chế độ `60s sáng` hoặc `60s tối`.
- Với `60s sáng`, thư mục input cần chứa `BT60SAM_*.xlsx` và các file `.rtf`.
- Với `60s tối`, thư mục input cần chứa `BT60SCHIEU_*.xlsx` và các file `.rtf`.
- Chọn thư mục output, hoặc để trống để app tự tạo thư mục `output` trong input.
- Nhập `Mã bản tin $a090`. Ô này chỉ gợi ý mờ `K303324`, không tự dùng làm giá trị.
- Kiểm tra hoặc sửa `Người biên mục $a911`. Mặc định bản sáng là `Trần Ngọc Thanh Hiền`, bản tối là `Lê Thị Mai Liên`.
- Bấm `Bắt đầu biên mục`.

## Output

App tạo file:

```text
Map_BanTin60GiaySang_{YYYY}_ Thang{MMDD}.xlsx
Map_BanTin60GiayChieu_{YYYY}_ Thang{MMDD}.xlsx
```

Ví dụ:

```text
Map_BanTin60GiaySang_2026_ Thang0604.xlsx
Map_BanTin60GiayChieu_2026_ Thang0604.xlsx
```

Các cột trong file Map:

- `$a090`: mã bản tin người dùng nhập.
- `$a500`: ê-kíp sản xuất.
- `$a505`: danh sách tin theo thứ tự playlist.
- `$a911`: người biên mục, chỉ ghi ở dòng dữ liệu đầu tiên.

## Log

App ghi log vào `app_bien_muc_60s.log` trong cùng thư mục chương trình. Khi mở app, log của các ngày cũ được xóa để file log không phình to theo thời gian.

## Quy tắc chính

Quy tắc bắt dòng tin dùng chung cho cả `60s sáng` và `60s tối`:

- Cột A được trim và không phân biệt chữ hoa/thường.
- Cột A phải bắt đầu bằng một trong các prefix: `60s`, `gat60s `, `60st`, hoặc `live -`.
- Không bắt nếu Cột A bắt đầu bằng `60s W `.
- Bỏ qua nếu Cột A chứa `coming up`, `nhung nguoi thuc hien`, hoặc ` end`.
- Cột C bắt buộc có ID dạng số. Nếu Cột C trống hoặc bắt đầu bằng chữ, ví dụ `qc123`, dòng đó không được bắt.

Tiêu đề tin được tách từ phần đầu file RTF theo thứ tự xuất hiện:

- Lấy dòng đầu tiên IN HOA, màu GREEN, dài hơn 20 ký tự và thuộc phần đầu kịch bản. Dòng này không bắt buộc BOLD.
- Nếu dòng ngay phía dưới cũng IN HOA, GREEN và dài hơn 20 ký tự thì nối thêm vào tiêu đề.
- Tiêu đề chỉ gồm 1 hoặc 2 dòng, không nối nhiều hơn.
- Dừng vùng tìm tiêu đề khi gặp dòng thân bài màu đỏ hoặc màu đen, dòng bắt đầu bằng `PB `, hoặc separator `===`.
- Loại tuyệt đối các dòng cue bắt đầu bằng `C1`, `C2`, `C3`, `C4`; không lấy các dòng này làm tiêu đề.

Ê-kíp sản xuất được đọc từ file có tên chứa `NHUNG NGUOI THUC HIEN.rtf`. Nếu thiếu file hoặc thiếu chức danh, app bổ sung từ tên file RTF tiền tố như `BGĐ`, `BPT`, `BT`, `BD`, `MC`, `ĐD`, `KT`.
