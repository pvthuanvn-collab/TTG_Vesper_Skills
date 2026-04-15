---
name: "sap-electricity-import"
description: "Chuyển file hóa đơn điện Excel thành bộ file import SAP JE theo template tham chiếu người dùng đã chốt."
alwaysAllow:
  - Bash
  - Write
---

# SAP Electricity Import

Skill này dùng để tạo bộ file import SAP JE từ file hóa đơn điện Excel.

## SOP thực hiện
1. Xác định file đầu vào là Excel hóa đơn điện hoặc thư mục chứa file đó.
2. Kiểm tra sheet dữ liệu và header thực tế của file nguồn.
3. Skill dò `header row` và map cột theo `header-based mapping + alias`, không phụ thuộc cứng vào vị trí cột.
4. Áp dụng rule chuyển đổi theo template JE mới đã được người dùng xác nhận.
5. Sinh bộ output review-ready trong thư mục `output_<MM>_<YYYY>`, không tự động post vào SAP.
6. Đối chiếu tổng tiền đầu vào với tổng Debit và Credit đầu ra.
7. Báo lại profile input detect được, giả định, control points và exception nếu có.

## Input mapping hiện tại
Skill hỗ trợ đọc dữ liệu theo tên cột và alias, ưu tiên các cột logic sau:
- `row_no`
- `gross_amount`
- `invoice_no`
- `invoice_series`
- `issue_date`

### Các alias chính đang hỗ trợ
- `row_no`
  - `STT`
- `gross_amount`
  - `TONG_NO`
  - `Tổng nợ`
  - `Tổng tiền`
  - `Gross Amount`
- `invoice_no`
  - `số HD`
  - `Số HĐ`
  - `SO HD`
  - `Sery HĐ`
  - `Số hóa đơn`
  - `Invoice No`
- `invoice_series`
  - `Ký hiệu`
  - `Seri HĐ`
  - `Mã kí hiệu`
  - `Mã ký hiệu`
  - `Invoice Series`
- `issue_date`
  - `NGÀY PHÁT HÀNH`
  - `Ngày PH`
  - `Ngày phát hành`
  - `Issue Date`

### Profile input hiện có
- `legacy_layout`
  - ví dụ: `số HD`, `Ký hiệu`, `NGÀY PHÁT HÀNH`
- `evn_layout_202603`
  - ví dụ: `Sery HĐ`, `Mã kí hiệu`, `Ngày PH`
- `generic_header_mapping`
  - fallback khi detect được cột bắt buộc nhưng không match profile signature cụ thể

## Rule chuyển đổi hiện tại
Các rule dưới đây bám theo template JE mới đã confirm:
- Toàn bộ file đầu vào chỉ sinh **1 journal / 1 dòng header** trên sheet `JE-Header` cho tất cả hóa đơn.
- Mỗi hóa đơn vẫn sinh **3 dòng line** trên sheet `JE-Line` và cùng thuộc về journal duy nhất đó:
  - 1 dòng chi phí `62721001` (Debit)
  - 1 dòng VAT `13311001` (Debit)
  - 1 dòng công nợ `33111001` (Credit)
- VAT mặc định: `8%`
- Base amount = `ROUND(TONG_NO / 1.08)`
- VAT amount = `TONG_NO - Base amount`
- Vendor mặc định:
  - `Mã đối tác`: `V00000162`
  - `Tên đối tác`: `CHI NHÁNH TỔNG CÔNG TY ĐIỆN LỰC TPHCM TNHH-CÔNG TY ĐIỆN LỰC SÀI GÒN`
  - `MST`: `0300951119-001`
- Project: `M02`
- Voucher Type: `7012`
- Branch / Note for Import: `7`
- Costing split dòng chi phí giữ như rule cũ, tách từ:
  - `12090310;M999994;M02;PMO;M0100000`
- Tax Group dòng VAT: `PVN5`
- Diễn giải mặc định suy ra từ ngày phát hành, mặc định lấy kỳ tiêu thụ là **tháng trước của ngày phát hành**.
- Sheet `JE-Header` cột `D` (`Memo`) dùng format rút gọn:
  - `Điện tiêu thụ tháng <m> năm <yyyy>`
  - bỏ phần `từ ngày ... đến ngày ...`
- Với journal header duy nhất:
  - `ReferenceDate`, `TaxDate` lấy theo **ngày tạo file import**.
  - `Memo` lấy theo **ngày phát hành sau cùng** trong danh sách hóa đơn.
- Sheet `JE-Line`:
  - cột `I` (`DueDate`) lấy theo `ReferenceDate` trên sheet `JE-Header`.
  - cột `L` (`LineMemo`) lấy đúng theo giá trị cột `D` của sheet `JE-Header`.
  - cột `M` (`ReferenceDate1`) lấy theo `ReferenceDate` trên sheet `JE-Header`.
  - thêm cột `TaxDate` trong template và lấy theo **ngày thực tế của hóa đơn**.

## Output template mới
Skill sinh theo template đã được đóng gói sẵn ngay trong skill:
- Template workbook nội bộ: `templates/SAP Import JE bằng WB.xlsx`
- Workbook output: `JE-Header`, `JE-Line`
- Text export:
  - `Header.txt`
  - `Line.txt`

## Validation hiện tại
### Validation đầu vào
- Skill quét tối đa `20` dòng đầu để tìm `header row`.
- Chỉ chạy tiếp nếu detect được đủ các cột logic bắt buộc:
  - `row_no`
  - `gross_amount`
  - `invoice_no`
  - `invoice_series`
  - `issue_date`
- Chặn các dòng thiếu:
  - `gross_amount`
  - `invoice_no`
  - `invoice_series`
  - `issue_date`

### Validation đầu ra
- Tổng `TONG_NO` đầu vào phải bằng tổng Debit đầu ra.
- Tổng `TONG_NO` đầu vào phải bằng tổng Credit đầu ra.
- Sheet `JE-Header` chỉ có `1` dòng journal cho toàn bộ hóa đơn.
- Số dòng output trên `JE-Line` phải bằng `3 x số hóa đơn`.
- `LineNum` trên `JE-Line` phải chạy tuần tự trong journal duy nhất.
- Kiểm tra file output chính, `Header.txt`, `Line.txt`, `.summary.json` đều được sinh ra.

### Summary output
Sau khi chạy, skill sinh thêm file `.summary.json` để phục vụ review với các thông tin như:
- `invoice_count`
- `header_row_count`
- `line_row_count`
- `input_total`
- `debit_total`
- `credit_total`
- `sheet_name`
- `header_row`
- `input_profile`
- `detected_headers`
- `skipped_rows`
- `output_folder`
- `period_month`
- `period_year`

## Cách chạy
### Truyền trực tiếp file đầu vào
```bash
python "$VESPER_SKILL_DIR/scripts/build_sap_electricity_import.py" "C:/path/input.xlsx"
```

### Truyền thư mục chứa file đầu vào
```bash
python "$VESPER_SKILL_DIR/scripts/build_sap_electricity_import.py" "C:/path/folder"
```

### Chỉ định file output
```bash
python "$VESPER_SKILL_DIR/scripts/build_sap_electricity_import.py" "C:/path/input.xlsx" "C:/path/output/target.xlsx"
```

## Output
Mặc định skill tạo thư mục con theo kỳ tiêu thụ:
- `output_<MM>_<YYYY>/`

Trong đó sinh ra:
- `SAP_Import by JE_<MM>_<YYYY>.xlsx`
- `HOA DON TIEN DIEN_ SAP_<sheetname>.summary.json`
- `Header.txt`
- `Line.txt`

## Control points cần kế toán confirm
- Xác nhận lại VAT, tài khoản, vendor và các field fix trước khi dùng rộng.
- Với format EVN mới, cần confirm mapping nghiệp vụ:
  - `Sery HĐ` có đúng là **Số HĐ** hay không
  - `Mã kí hiệu` / `Mã ký hiệu` có đúng là **Seri/Ký hiệu HĐ** hay không
- Confirm `TaxGroup`, `VoucherType`, `AP account`, costing split là đúng với template SAP đang dùng.
- Confirm `ReferenceDate`, `TaxDate`, `DueDate` cùng lấy theo `issue_date` là phù hợp.

## Assumptions / Risks
- Rule hiện tại vẫn chuyên biệt cho bộ hóa đơn điện mẫu đã cung cấp.
- VAT, account, vendor, project, costing split, voucher type đang hard-coded theo xác nhận hiện tại.
- Script linh hoạt hơn ở phần đọc input, nhưng vẫn giả định sheet dữ liệu nằm ở sheet đầu tiên của workbook.
- Template workbook đã được bundle ngay trong skill tại `templates/SAP Import JE bằng WB.xlsx`. Khi copy skill cho user khác, cần copy nguyên cả folder skill gồm thư mục `templates/`.
- Nếu template nội bộ thay đổi cấu trúc, cần cập nhật lại skill tương ứng.
- Nếu file nguồn xuất hiện header mới chưa có trong alias list thì cần bổ sung alias hoặc profile input tương ứng.
