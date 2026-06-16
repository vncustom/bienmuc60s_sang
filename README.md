# App biên mục 60s sáng

Ứng dụng tạo file Map tổng cho bản tin 60s sáng từ thư mục input gồm playlist Excel và các file kịch bản RTF.

## Cài đặt

```powershell
pip install -r requirements.txt
```

## Chạy app

```powershell
python app_bien_muc_60s.py
```

Trên giao diện:

- Chọn thư mục input chứa `BT60SAM_*.xlsx` và các file `.rtf`.
- Chọn thư mục output, hoặc để trống để app tự tạo thư mục `output` trong input.
- Nhập `Mã bản tin $a090`. Ô này chỉ gợi ý mờ `K303324`, không tự dùng làm giá trị.
- Kiểm tra hoặc sửa `Người biên mục $a911`.
- Bấm `Bắt đầu biên mục`.

## Output

App tạo file:

```text
Map_BanTin60GiaySang_{YYYY}_ Thang{MMDD}.xlsx
```

Ví dụ:

```text
Map_BanTin60GiaySang_2026_ Thang0604.xlsx
```

Các cột trong file Map:

- `$a090`: mã bản tin người dùng nhập.
- `$a500`: ê-kíp sản xuất.
- `$a505`: danh sách tin theo thứ tự playlist.
- `$a911`: người biên mục, chỉ ghi ở dòng dữ liệu đầu tiên.

## Log

App ghi log vào `app_bien_muc_60s.log` trong cùng thư mục chương trình. Khi mở app, log của các ngày cũ được xóa để file log không phình to theo thời gian.

## Quy tắc chính

App lọc tin từ playlist bằng các dòng bắt đầu `60sa ` hoặc `Live - 60sa`, có ID số và trạng thái `ONLINE`. Các mục như `W60sa`, quảng cáo, coming up, gạt, hình hiệu không được đưa vào Map.

Tiêu đề tin được tách từ RTF theo quy tắc chung: dòng đầu tiên IN HOA, BOLD, màu GREEN, dài hơn 15 ký tự. Nếu dòng ngay dưới cũng IN HOA, BOLD, GREEN và hợp lệ thì gộp thêm, tối đa 2 dòng. Các dòng cue hình như `CẬN GIỮA`, `TOÀN GIỮA`, `CẬN PHẢI` không được coi là tiêu đề.

Ê-kíp sản xuất được đọc từ file có tên chứa `NHUNG NGUOI THUC HIEN.rtf`. Nếu thiếu file hoặc thiếu chức danh, app bổ sung từ tên file RTF tiền tố như `BGĐ`, `BPT`, `BT`, `BD`, `MC`, `ĐD`, `KT`.
