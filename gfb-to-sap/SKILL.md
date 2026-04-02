---
name: gfb-to-sap
description: >
  Chuyển đổi báo cáo cước Grab For Business (GFB Billing Calculation Report)
  thành file Excel import SAP B1. Skill này đọc file GFB Excel, tổng hợp
  chi phí theo phòng ban (GROUP_NAME), và xuất file SAP Import với 2 dòng
  mỗi phòng ban: dòng chi phí (Debit TK chi phí) + dòng VAT (Debit TK thuế
  GTGT đầu vào). Kích hoạt khi nghe thấy: "GFB", "Grab For Business",
  "hạch toán Grab", "chi phí Grab", "cước Grab", "import SAP Grab",
  "bút toán Grab", "chuyển đổi GFB", hoặc bất kỳ yêu cầu nào liên quan
  đến xử lý hóa đơn Grab và nhập vào SAP.
---

# Skill: GFB Billing → SAP Import

## Mục đích
Tự động chuyển đổi file báo cáo cước Grab For Business (GFB Billing Calculation Report)
thành file Excel theo đúng format SAP Import Template, **tổng hợp theo phòng ban**.

## Thông tin cố định (đã cấu hình)
| Thông số | Giá trị |
|---|---|
| Nhà cung cấp (Mã đối tác) | GRAB |
| Tên đối tác | CÔNG TY TNHH GRAB |
| MST Grab | 0312650437 |
| TK chi phí (GL) | 111111111 |
| TK VAT đầu vào (GL) | 22222222 |
| Distr. Rule | A00001 |
| Nhóm thuế | PVN5 |
| Branch | LEGACY |
| Tình trạng kê khai | Kê khai |

## Cấu trúc output
Mỗi phòng ban tạo **2 dòng** trong file SAP:
- **Dòng 1 – Chi phí**: Debit TK `111111111`, Amount = PRE_VAT (vận chuyển + dịch vụ)
- **Dòng 2 – VAT**: Debit TK `22222222`, Amount = VAT (8%), Base Amount = PRE_VAT

Ngày hạch toán = **cuối tháng chiếm đa số** của VAT_INVOICE_DATE trong file GFB
(ví dụ: file có 173/180 hóa đơn tháng 02 → ngày hạch toán = 28/02/2026).

## Quy trình thực hiện

### Bước 1: Xác định file đầu vào
Tìm file GFB Billing trong thư mục người dùng chỉ định. File thường có tên dạng:
`GFB Billing Calculation Report_*.xlsx`

Nếu người dùng không chỉ rõ, hỏi: "Bạn vui lòng cho biết đường dẫn đến file GFB Billing?"

### Bước 2: Xác định file đầu ra
Tên file output mặc định: `GFB_SAP_Import_T{MM}_{YYYY}.xlsx`
Lưu vào cùng thư mục với file input (hoặc theo yêu cầu người dùng).

### Bước 3: Chạy script chuyển đổi
Sử dụng script tại `scripts/convert_gfb_to_sap.py`:

```bash
python3 scripts/convert_gfb_to_sap.py \
  "<đường dẫn file GFB>" \
  "<đường dẫn file output>" \
  --expense-gl 111111111 \
  --vat-gl 22222222 \
  --vendor-code GRAB \
  --distr-rule A00001 \
  --tax-group PVN5
```

**Tùy chọn thêm** (nếu người dùng muốn override):
- `--posting-date DD/MM/YYYY` — chỉ định ngày hạch toán cụ thể
- `--expense-gl <số TK>` — thay đổi TK chi phí
- `--vat-gl <số TK>` — thay đổi TK VAT
- `--distr-rule <mã>` — thay đổi Distr. Rule

### Bước 4: Báo cáo kết quả
Sau khi chạy xong, trình bày cho người dùng:
1. Bảng tổng hợp theo phòng ban (số chuyến, chi phí, VAT, tổng)
2. Ngày hạch toán được xác định
3. Số dòng SAP đã tạo
4. Link để mở file output

### Bước 5: Trình bày file output
Dùng `present_files` hoặc cung cấp link `computer://` đến file output.

## Xử lý các tình huống đặc biệt

**File GFB có nhiều tháng**: Script tự động lấy tháng chiếm đa số.
Nếu cần tách riêng từng tháng, hỏi người dùng trước khi chạy.

**Phòng ban không xác định (GROUP_NAME = null)**: Được gộp vào nhóm "Unknown".
Nhắc người dùng kiểm tra và cập nhật thủ công nếu cần.

**Tổng tiền không khớp**: Kiểm tra các chuyến có AMOUNT = 0 (bị loại bỏ tự động).
Thông báo số chuyến bị loại bỏ cho người dùng.

**Người dùng muốn thay đổi mapping tài khoản**: Chấp nhận và truyền qua tham số
`--expense-gl`, `--vat-gl`, v.v. Không cần sửa script.

## Kiểm tra kết quả (validation)
Sau khi tạo file, tự động kiểm tra:
- Tổng Debit TK chi phí + Tổng Debit TK VAT = Tổng AMOUNT trong GFB
- Số phòng ban trong output = số GROUP_NAME unique trong input
- File output có sheet tên "HOA DON" với đúng 40 cột

Nếu có sai lệch, báo ngay cho người dùng trước khi trình bày file.

## Ví dụ diễn giải (Remarks/Diễn giải)
```
Chi phí đi lại Grab tháng 02/2026 - TCKT
Chi phí đi lại Grab tháng 02/2026 - ADM
Chi phí đi lại Grab tháng 02/2026 - General
```
