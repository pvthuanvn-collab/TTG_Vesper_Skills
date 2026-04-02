---
name: "SKILL: Accounting: Chuyển đổi hóa đơn điện -→ SAP (Excel)"
description: "Chuyển file hóa đơn tiền điện từ Excel nguồn (điện lực) sang file Excel theo mẫu import SAP (2 dòng/hoá đơn: chi phí + thuế)."
globs: ["*.xlsx", "*.xls"]
alwaysAllow: ["Bash"]
---

# Mục tiêu
Khi user gọi **[skill:chuyen-doi-hoa-don-dien-sap]** và đính kèm **file Excel nguồn** (danh sách hoá đơn điện), bạn sẽ tạo **file Excel output** theo đúng format import SAP.

- **1 hoá đơn gốc → 2 dòng SAP**
  - Dòng 1: Chi phí **TK 62721001**
  - Dòng 2: Thuế GTGT **TK 13331001**
- VAT mặc định **8%**
- Sắp xếp theo **Số HĐ** tăng dần

# Input yêu cầu
File Excel nguồn phải có tối thiểu các cột (đúng tên cột):
- `số HD`
- `Ký hiệu`
- `TONG_NO`
- `NGÀY PHÁT HÀNH`

# Hành vi cần thực hiện
## 1) Nhận file nguồn
- Nếu user đính kèm **nhiều file**, hỏi user chọn file nào.
- Lấy **đường dẫn file** (path) của attachment làm input.

## 2) Tự xác định tháng/năm kỳ hoá đơn
- Ưu tiên **tự suy ra** từ cột `NGÀY PHÁT HÀNH`.
- Nếu file có **nhiều tháng/năm khác nhau** và user chưa chỉ định:
  - Hỏi user muốn xuất theo **tháng/năm nào**.

## 3) Chạy chuyển đổi
Chạy script converter đi kèm skill:

- Mặc định (tự suy tháng/năm, nếu dữ liệu chỉ có 1 tháng/năm):
```bash
python "C:\\Users\\thuan.pv\\.sophie-agent\\workspaces\\my-workspace\\skills\\chuyen-doi-hoa-don-dien-sap\\converter.py" "<INPUT_XLSX>" --output "<SESSION_DATA>\\HOA_DON_DIEN_SAP.xlsx"
```

- Nếu user chỉ định tháng/năm:
```bash
python "C:\\Users\\thuan.pv\\.sophie-agent\\workspaces\\my-workspace\\skills\\chuyen-doi-hoa-don-dien-sap\\converter.py" "<INPUT_XLSX>" --month <MM> --year <YYYY> --output "<SESSION_DATA>\\HOA_DON_DIEN_SAP_T<MM><YY>.xlsx"
```

**Quan trọng**:
- Output luôn ghi vào **session data folder** của cuộc hội thoại hiện tại.
- Làm tròn chi phí theo kiểu **Excel ROUND (half-up)**.

## 4) Trả kết quả cho user
- Trả về:
  - Đường dẫn file output (clickable)
  - Tóm tắt: số hoá đơn input, số dòng SAP output
  - Nếu script báo dữ liệu có nhiều tháng/năm → hỏi user chọn kỳ.

# Kiểm tra chất lượng (bắt buộc)
Sau khi xuất file:
- Số dòng output = 2 × số hoá đơn hợp lệ
- Mỗi `Số HĐ` có đúng 2 dòng
- Tổng `Debit` (2 dòng) = `TONG_NO` gốc

# Ví dụ hội thoại
User: "[skill:chuyen-doi-hoa-don-dien-sap] Đây là file DS_HD_DIEN.xlsx"
Assistant:
1) Chạy converter
2) Trả file `HOA_DON_DIEN_SAP_T1125.xlsx` kèm tóm tắt
