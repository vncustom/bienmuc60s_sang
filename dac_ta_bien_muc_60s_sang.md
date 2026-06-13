# Đặc tả biên mục bản tin 60s sáng

## Mục tiêu

Tự động tạo file Map tổng cho bản tin 60s sáng từ một thư mục input gồm:

- File playlist Excel dạng `BT60SAM_YYYYMMDD.xlsx`.
- Các file kịch bản `.rtf` tương ứng với từng mục trong playlist.

Kết quả cần giống file mẫu:

`Map_BanTin60GiaySang_2026_ Thang0604.xlsx`

## File đầu vào

### Playlist Excel

File mẫu: `BT60SAM_20260604.xlsx`.

Sheet đầu tiên chứa toàn bộ rundown. Các cột dùng trong biên mục:

| Cột | Nội dung | Cách dùng |
| --- | --- | --- |
| A | Tên mục / tên file | Dùng để lọc tin và tìm file RTF |
| C | ID video | Ghi vào cuối `$a505` |
| D | Trạng thái | Chỉ lấy mục `ONLINE` |
| F | Thời lượng phát | Chuyển thành `00:mm:ss` |

Ngày phát sóng lấy từ dòng đầu cột A, ví dụ:

`CHUONG TRINH 60 GIAY SANG NGAY 04-06-2026`

### File RTF

Tên file RTF thường trùng hoặc gần trùng với cột A trong playlist. Khi so khớp cần bỏ qua khác biệt do dấu `/`, khoảng trắng và ký tự đặc biệt.

Ví dụ:

- Playlist: `60sa 030626 BRITAIN-BOE/BANKNOTES`
- File RTF: `60sa 030626 BRITAIN-BOEBANKNOTES.rtf`

## Quy tắc lọc tin

Chỉ lấy các dòng thỏa tất cả điều kiện:

1. Cột A bắt đầu bằng `60sa ` hoặc `Live - 60sa`.
2. Không lấy các mục bắt đầu bằng `W60sa`.
3. Cột C là ID dạng số.
4. Cột D là `ONLINE`.

Các mục bị loại:

- `HINH HIEU`
- `CHAO DAU & HEADLINES`
- `COMING UP`
- `QUANG CAO`
- `GAT KINH TE`
- `GAT VAN HOA`
- Các dòng `W60sa ...`
- `CHAO CUOI`
- `60s END`

Với input mẫu, kết quả là 24 tin.

## Quy tắc lấy thời lượng

Thời lượng lấy từ cột F của playlist, không dùng cột E hoặc G.

Trong playlist mẫu, Excel hiển thị thời lượng kiểu `00:52:00` nhưng ý nghĩa thực tế là `00 phút 52 giây`. Khi đọc bằng `openpyxl`, giá trị này thường thành `datetime.time(0, 52)`, cần hiểu trường `hour` là phút và trường `minute` là giây.

Giá trị `datetime.time(0, 52)` được đổi thành:

`00:00:52`

Nếu giá trị là chuỗi hoặc số serial Excel thì cũng cần quy đổi về cùng format `00:mm:ss`.

## Quy tắc lấy tiêu đề

Tiêu đề lấy từ định dạng RTF, không lấy theo ID cố định.

Quy tắc chung:

1. Tìm dòng đầu tiên trong RTF thỏa đồng thời:
   - IN HOA.
   - BOLD.
   - Màu GREEN trong bảng màu RTF, thường là `rgb(0,128,0)`.
   - Dài hơn 15 ký tự.
2. Nếu ngay dưới dòng tiêu đề có thêm một dòng cũng IN HOA, BOLD, GREEN và hợp lệ thì gộp 2 dòng bằng một dấu cách.
3. Chỉ gộp tối đa 2 dòng, không gộp dòng thứ 3.
4. Không coi là tiêu đề nếu dòng có một trong các cụm kỹ thuật:
   - `CẬN TRÁI`
   - `CẬN GIỮA`
   - `CẬN PHẢI`
   - `TRUNG GIỮA`
   - `TOÀN PHẢI`
   - `TOÀN TRÁI`
   - `TOÀN GIỮA`
5. Không nối thêm `Sapo:` vào tiêu đề; `Sapo:` là nội dung riêng, không phải quy tắc chung.

Ví dụ trong `60sa 040626 EU-TECH.rtf`, tiêu đề có 2 dòng BOLD GREEN:

```
NGHỊ VIỆN CHÂU ÂU ĐỔI CÔNG CỤ TÌM KIẾM MẶC ĐỊNH,
GIẢM PHỤ THUỘC VÀO CÔNG NGHỆ MỸ
```

Kết quả đưa vào `$a505`:

`NGHỊ VIỆN CHÂU ÂU ĐỔI CÔNG CỤ TÌM KIẾM MẶC ĐỊNH, GIẢM PHỤ THUỘC VÀO CÔNG NGHỆ MỸ`

## File output

Tên file:

`Map_BanTin60GiaySang_{YYYY}_ Thang{MM}{DD}.xlsx`

Ví dụ:

`Map_BanTin60GiaySang_2026_ Thang0604.xlsx`

Sheet: `Sheet1`.

Header:

| A | B | C | D |
| --- | --- | --- | --- |
| `$a090` | `$a500` | `$a505` | `$a911` |

## Quy tắc ghi dữ liệu

### Cột `$a090`

Mã bản tin, do người dùng nhập trên app. Với mẫu hiện tại:

`K303324`

Mã này lặp lại ở tất cả dòng dữ liệu, kể cả dòng trống cuối.

### Cột `$a500`

Bóc tách thông tin ê-kíp sản xuất, ưu tiên từ file `NHUNG NGUOI THUC HIEN.rtf` hoặc file có tên chứa cụm `NHUNG NGUOI THUC HIEN`.

File RTF ê-kíp có thể có nhãn chức danh trên một dòng và tên ở các dòng sau. App gom toàn bộ tên dưới cùng một chức danh cho đến khi gặp chức danh kế tiếp, rồi nối bằng dấu phẩy.

Các dòng `$a500` theo thứ tự cố định:

| Dòng | Format |
| --- | --- |
| 1 | `Ban Giám đốc: {tên}` |
| 2 | `Biên tập: {tên}` |
| 3 | `Biên dịch: {tên}` |
| 4 | `Dẫn chương trình: {tên}` |
| 5 | `Đạo diễn: {tên}` |
| 6 | `Kỹ thuật: {tên}` |
| 7 | `Quay phim: {tên}` |

Ví dụ:

```text
Ban Giám đốc: TRẦN VĂN KHÁNH
Biên tập: MAI LAN, NHẬT MINH, BẢO DUY
Biên dịch: MINH ĐỨC, NHƯ YẾN
Dẫn chương trình: HOÀNG VĨNH, MINH NGỌC
Đạo diễn: HOÀN THIỆN
Kỹ thuật: HỒNG VÂN, TRUNG THÀNH, QUANG HUY
Quay phim: TRẦN TÚ
```

Nếu không có file `NHUNG NGUOI THUC HIEN.rtf`, hoặc chức danh trong file này bị thiếu tên, app bổ sung từ tên các file RTF tiền tố:

| Tiền tố file | Gán vào |
| --- | --- |
| `BGĐ ` hoặc `BPT ` | `Ban Giám đốc` |
| `BT ` | `Biên tập` |
| `BD ` | `Biên dịch` |
| `MC ` | `Dẫn chương trình` |
| `ĐD ` | `Đạo diễn` |
| `KT ` | `Kỹ thuật` |

Chức danh nào không có tên thì vẫn ghi dòng chức danh, phần sau dấu `:` để trống.

### Cột `$a505`

Mỗi tin một dòng theo format:

`{STT} - {TIÊU ĐỀ}. Thời lượng: {00:mm:ss}. ID: {ID}`

STT có 2 chữ số, bắt đầu từ `01`.

Ví dụ:

`01 - TP. HỒ CHÍ MINH: ĐIỂM CHUẨN LỚP 10 CÓ THỂ TĂNG 0,25-2 ĐIỂM. Thời lượng: 00:00:52. ID: 260603212`

### Cột `$a911`

Chỉ điền ở dòng dữ liệu đầu tiên:

`Trần Ngọc Thanh Hiền`

Các dòng sau để trống.

### Dòng trống cuối

Sau 24 dòng tin, file mẫu có thêm một dòng chỉ có `$a090`, các cột còn lại trống. App cần tạo dòng này để khớp output mẫu.

## Định dạng Excel

Để khớp file mẫu:

- Font header và cột A/D: Times New Roman 13.
- Cột C dữ liệu: Times New Roman 14.
- Độ rộng cột:
  - A: `10.6640625`
  - B: `48.6640625`
  - C: `77.21875`
  - D: `9.109375`

## Luồng xử lý

1. Người dùng chọn thư mục input.
2. Người dùng chọn thư mục output, nếu bỏ trống thì tạo `output` trong thư mục input.
3. App tìm file `BT60SAM_*.xlsx`.
4. Đọc ngày phát sóng từ dòng đầu hoặc từ tên file.
5. Lọc danh sách 24 tin theo quy tắc 60s sáng.
6. Với từng tin, tìm RTF tương ứng và trích tiêu đề.
7. Chuyển thời lượng từ cột F sang `00:mm:ss`.
8. Ghi file Map theo đúng format.
9. Báo đường dẫn file đã tạo.
