/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from "react";
import { FileCode2, Copy, Check, Info, Download } from "lucide-react";

export default function PythonScriptView() {
  const [copied, setCopied] = useState(false);

  const pythonCode = `import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import re
import datetime

def process_manifest(input_path_or_df):
    """
    Hàm đọc dữ liệu Manifest gốc (tương đương Hình 1), tự động hóa các bước mapping:
    - Cột ROUTE: SGN-HAN-ICN nếu MAWB NO chứa 'HAN', ngược lại SGN-ICN
    - Cột WEIGHT: Lấy từ cột 'R.W/T'
    - Cột Handling charge: Cố định 10.00 USD
    - Cột Warehouse charge: Lấy hoặc tham chiếu từ cột 'HAWB'
    """
    # 1. Đọc file đầu vào (Hỗ trợ truyền file Excel/CSV hoặc trực tiếp Pandas DataFrame)
    if isinstance(input_path_or_df, str):
        # Nếu đường dẫn kết thúc bằng .xlsx hoặc .xls, dùng read_excel, ngược lại read_csv
        if input_path_or_df.endswith(('.xlsx', '.xls')):
            df_raw = pd.read_excel(input_path_or_df)
        else:
            df_raw = pd.read_csv(input_path_or_df)
    else:
        df_raw = input_path_or_df.copy()

    processed_rows = []
    
    for idx, row in df_raw.iterrows():
        # Lấy thông tin MAWB NO
        mawb_no = str(row.get('MAWB NO', '')).strip()
        
        # a. QUY TẮC PHÂN TÍCH ROUTE: Chứa 'HAN' -> 'SGN-HAN-ICN', ngược lại -> 'SGN-ICN'. Nếu MAWB không chứa chữ thì mặc định 'SGN-ICN'.
        has_letters = any(c.isalpha() for c in mawb_no)
        if has_letters and 'HAN' in mawb_no.upper():
            route = "SGN-HAN-ICN"
        else:
            route = "SGN-ICN"
            
        # b. QUY TẮC LẤY CHỈ SỐ WEIGHT: Lấy chính xác giá trị từ cột R.W/T
        weight = float(row.get('R.W/T', 0.0))
        
        # c. QUY TẮC HANDLING CHARGE: Cố định là 10.00 USD cho tất cả các dòng
        handling_charge = 10.00
        
        # e. QUY TẮC WAREHOUSE CHARGE: Lấy hoặc tham chiếu trực tiếp từ cột HAWB ở Hình 1
        # Trích xuất số cuối cùng từ chuỗi HAWB (Ví dụ: "HW-WH-45" -> 45.0) làm phí lưu kho bãi
        hawb_val = str(row.get('HAWB', ''))
        numbers_in_hawb = re.findall(r'\\d+', hawb_val)
        if numbers_in_hawb:
            warehouse_charge = float(numbers_in_hawb[-1])
        else:
            warehouse_charge = 15.0  # Giá trị mặc định nếu ko tìm thấy số
            
        processed_rows.append({
            'MAWB NO': mawb_no,
            'HAWB': hawb_val,
            'SHIPPER': str(row.get('SHIPPER', 'Unknown')).strip(),
            'CONSIGNEE': str(row.get('CONSIGNEE', 'Unknown')).strip(),
            'ROUTE': route,
            'WEIGHT': weight,
            'HANDLING_CHARGE': handling_charge,
            'WAREHOUSE_CHARGE': warehouse_charge,
            'OTH_CHG': 0.0  # Phí khác mặc định là 0.00
        })
        
    return pd.DataFrame(processed_rows)


def export_shipper_report(df_processed, shipper_name, unit_prices_dict, exchange_rate, output_path, month=6, year=2026):
    """
    Hàm lọc dữ liệu theo Shipper, áp đơn giá động, tự động tính tổng tiền USD và VND,
    và xuất ra file Excel được định dạng y hệt Hình 2 sử dụng openpyxl.
    """
    # Lọc dữ liệu theo tên Shipper chỉ định
    df_shipper = df_processed[df_processed['SHIPPER'].str.upper() == shipper_name.upper()].copy()
    
    if df_shipper.empty:
        print(f"Cảnh báo: Không có dữ liệu cho Shipper '{shipper_name}'")
        return False
        
    # Tạo một Workbook và Sheet mới
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bảng kê tính tiền"
    ws.views.sheetView[0].showGridLines = True  # Đảm bảo hiển thị lưới Excel
    
    # 1. THIẾT LẬP THÀNH PHẦN TIÊU ĐỀ DOANH NGHIỆP
    ws['A1'] = "GLOBAL EXPRESS LOGISTICS CO., LTD."
    ws['A1'].font = Font(name="Arial", size=10, bold=True, color="4F4F4F")
    
    ws['A2'] = "SGN Cargo Terminal, Tan Son Nhat Airport, Ho Chi Minh City, Vietnam"
    ws['A2'].font = Font(name="Arial", size=8, italic=True, color="7F7F7F")
    
    # 2. TIÊU ĐỀ CHÍNH CỦA BÁO CÁO (YÊU CẦU: In hoa, In đậm, Căn giữa)
    ws.merge_cells("A4:K4")
    title_cell = ws['A4']
    title_cell.value = f"BẢNG KÊ HÓA ĐƠN THÁNG {str(month).zfill(2)} NĂM {year} - {shipper_name.upper()}"
    title_cell.font = Font(name="Arial", size=14, bold=True, color="0A3622")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[4].height = 30
    
    # 3. THÔNG TIN CO-ORDINATION VÀ TỶ GIÁ (Exchange rate)
    ws.merge_cells("A5:K5")
    rate_cell = ws['A5']
    rate_cell.value = f"Tỷ giá áp dụng thanh toán (Exchange Rate): {exchange_rate:,.0f} VND/USD"
    rate_cell.font = Font(name="Arial", size=10, italic=True, bold=True, color="1E3A8A")
    rate_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 20
    
    # 4. CHUẨN BỊ THÀNH PHẦN CỘT (Cột 1 đến 11)
    headers = [
        "NO.", "MAWB NO", "HAWB", "ROUTE", "WEIGHT (KG)", 
        "UNIT PRICE (USD)", "FREIGHT CHARGE (USD)", "HANDLING CHARGE (USD)", 
        "WAREHOUSE CHARGE (USD)", "OTH.CHG (USD)", "TOTAL (USD)"
    ]
    
    # Định dạng của Header (YÊU CẦU: Màu xanh mint background, border đầy đủ)
    mint_fill = PatternFill(start_color="D1EAE5", end_color="D1EAE5", fill_type="solid")
    header_font = Font(name="Arial", size=9, bold=True, color="0F5E3D")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    border_side_medium = Side(border_style="medium", color="0F5E3D")
    border_side_thin = Side(border_style="thin", color="B3D8D1")
    header_border = Border(top=border_side_medium, bottom=border_side_medium, left=border_side_thin, right=border_side_thin)
    
    header_row_idx = 7
    ws.row_dimensions[header_row_idx].height = 28
    
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=header_row_idx, column=col_idx)
        cell.value = h_text
        cell.font = header_font
        cell.fill = mint_fill
        cell.alignment = center_align
        cell.border = header_border
        
    # 5. ĐỔ DỮ LIỆU VÀO CÁC DÒNG
    thin_border_side = Side(border_style="thin", color="CCCCCC")
    data_border = Border(top=thin_border_side, bottom=thin_border_side, left=thin_border_side, right=thin_border_side)
    
    yellow_fill = PatternFill(start_color="FDF9C3", end_color="FDF9C3", fill_type="solid")
    gold_font = Font(name="Arial", size=9, bold=True, color="854D0E")
    
    start_row = 8
    for i, (_, row) in enumerate(df_shipper.iterrows()):
        current_row = start_row + i
        ws.row_dimensions[current_row].height = 20
        
        # d. ĐƠN GIÁ ĐỘNG (Lấy từ từ điển hoặc mặc định là 3.0 nếu chưa khai báo)
        # Sử dụng HAWB làm khóa định danh đơn giá
        hawb_key = row['HAWB']
        unit_price = unit_prices_dict.get(hawb_key, 3.0)
        
        # g. Tính toán các cột tài chính
        weight = row['WEIGHT']
        freight_charge = weight * unit_price
        handling_charge = row['HANDLING_CHARGE']
        warehouse_charge = row['WAREHOUSE_CHARGE']
        other_charge = row['OTH_CHG']
        
        # TOTAL = Freight charge + Handling charge + Warehouse charge + Other charge
        total_usd = freight_charge + handling_charge + warehouse_charge + other_charge
        
        row_values = [
            i + 1,                   # NO.
            row['MAWB NO'],         # MAWB NO
            row['HAWB'],            # HAWB
            row['ROUTE'],           # ROUTE
            weight,                 # WEIGHT
            unit_price,             # UNIT PRICE
            freight_charge,         # FREIGHT CHARGE
            handling_charge,        # HANDLING CHARGE
            warehouse_charge,       # WAREHOUSE CHARGE
            other_charge,           # OTH_CHG
            total_usd               # TOTAL (USD)
        ]
        
        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.value = val
            cell.font = Font(name="Arial", size=9)
            cell.border = data_border
            
            # Căn lề thích hợp
            if col_idx in [1, 2, 3, 4]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                
            # Trình bày định dạng số chuyên nghiệp
            if col_idx == 5:
                cell.number_format = '#,##0.0'       # Định dạng WEIGHT
            elif col_idx in [6, 7, 8, 9, 10]:
                cell.number_format = '$#,##0.00'     # Định dạng Phí USD
            elif col_idx == 11:
                cell.number_format = '$#,##0.00'     # Cột TOTAL (USD)
                cell.font = gold_font
                cell.fill = yellow_fill              # TO MÀU VÀNG CẢ CỘT TOTAL (USD) ĐỂ NỔI BẬT

    # 6. DÒNG TÍNH TỔNG (TOTAL FOOTER - Kẻ bảng viền đậm nền Mint)
    total_row_idx = start_row + len(df_shipper)
    ws.row_dimensions[total_row_idx].height = 24
    
    # Gộp A đến D làm chữ "TOTAL / TỔNG CỘNG"
    ws.merge_cells(start_row=total_row_idx, start_column=1, end_row=total_row_idx, end_column=4)
    total_label_cell = ws.cell(row=total_row_idx, column=1)
    total_label_cell.value = "TOTAL / TỔNG CỘNG"
    total_label_cell.font = Font(name="Arial", size=9, bold=True, color="0F5E3D")
    total_label_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Đắp border và fill màu mint cho vùng gộp
    for c in range(1, 5):
        cell = ws.cell(row=total_row_idx, column=c)
        cell.fill = mint_fill
        cell.border = Border(top=Side(style='medium', color="0F5E3D"), bottom=Side(style='medium', color="0F5E3D"), left=thin_border_side, right=thin_border_side)
        
    # Xuất công thức SUM tự động cho các cột còn lại
    col_letters = {5: 'E', 7: 'G', 8: 'H', 9: 'I', 10: 'J', 11: 'K'}
    
    # Cột đơn giá (Cột 6) để dấu "-"
    dash_cell = ws.cell(row=total_row_idx, column=6)
    dash_cell.value = "-"
    dash_cell.alignment = Alignment(horizontal="center", vertical="center")
    dash_cell.fill = mint_fill
    dash_cell.border = Border(top=Side(style='medium', color="0F5E3D"), bottom=Side(style='medium', color="0F5E3D"), left=thin_border_side, right=thin_border_side)
    
    for col_idx, letter in col_letters.items():
        cell = ws.cell(row=total_row_idx, column=col_idx)
        cell.value = f"=SUM({letter}{start_row}:{letter}{total_row_idx-1})"
        cell.font = Font(name="Arial", size=9, bold=True)
        cell.fill = mint_fill
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.border = Border(top=Side(style='medium', color="0F5E3D"), bottom=Side(style='medium', color="0F5E3D"), left=thin_border_side, right=thin_border_side)
        
        if col_idx == 5:
            cell.number_format = '#,##0.0'
        else:
            cell.number_format = '$#,##0.00'
            
        if col_idx == 11:
            # Nhấn mạnh tổng USD cuối cùng bằng màu vàng và màu đỏ đậm
            cell.fill = yellow_fill
            cell.font = Font(name="Arial", size=10, bold=True, color="990000")

    # 7. QUY ĐỔI TIỀN VND (YÊU CẦU: Tô màu vàng nổi bật, viền double dưới tổng kết)
    vnd_row_idx = total_row_idx + 2
    ws.row_dimensions[vnd_row_idx].height = 28
    
    ws.merge_cells(start_row=vnd_row_idx, start_column=1, end_row=vnd_row_idx, end_column=6)
    vnd_label = ws.cell(row=vnd_row_idx, column=1)
    vnd_label.value = "TỔNG TIỀN THANH TOÁN QUY ĐỔI (TOTAL VALUE IN VND):"
    vnd_label.font = Font(name="Arial", size=10, bold=True, color="1E3A8A")
    vnd_label.alignment = Alignment(horizontal="right", vertical="center")
    
    ws.merge_cells(start_row=vnd_row_idx, start_column=7, end_row=vnd_row_idx, end_column=11)
    vnd_val = ws.cell(row=vnd_row_idx, column=7)
    # Công thức quy đổi VND = Tổng USD (K{total_row_idx}) * Tỷ giá
    vnd_val.value = f"=K{total_row_idx}*{exchange_rate}"
    vnd_val.font = Font(name="Arial", size=11, bold=True, color="B91C1C")
    vnd_val.alignment = Alignment(horizontal="center", vertical="center")
    vnd_val.number_format = '#,##0" VND"'
    
    # Định dạng đường viền và màu nền vàng chanh nhạt cho dòng quy đổi VND
    double_bottom_border = Border(
        top=Side(style='medium', color="1E3A8A"), 
        bottom=Side(style='double', color="1E3A8A"),
        left=Side(style='thin', color="CCCCCC"),
        right=Side(style='thin', color="CCCCCC")
    )
    for col in range(1, 12):
        cell = ws.cell(row=vnd_row_idx, column=col)
        cell.fill = yellow_fill
        cell.border = double_bottom_border

    # Căn chỉnh độ rộng các cột tự động
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    # Lưu file kết quả
    wb.save(output_path)
    print(f"Báo cáo của shipper '{shipper_name}' đã được xuất thành công ra '{output_path}'")
    return True

# ==========================================
# CHƯƠNG TRÌNH CHẠY THỬ NGHIỆM ĐIỂN HÌNH (DEMO RUN)
# ==========================================
if __name__ == "__main__":
    print("--- Khởi chạy chương trình mô phỏng xử lý Manifest & Xuất báo cáo ---")
    
    # Giả lập dữ liệu thô tải lên (Tương ứng Hình 1)
    data_hinh_1 = {
        'MAWB NO': [
            "994-33069691 T/S HAN", 
            "TS HAN 994-38098642", 
            "180-18407734 TCS", 
            "180-18407734 TCS"
        ],
        'HAWB': ["HW-WH-45", "HW-WH-25", "HW-WH-80", "HW-WH-12"],
        'SHIPPER': [
            "DYM VIETNAM CO., LTD", 
            "DYM VIETNAM CO., LTD", 
            "SAMSUNG ELECTRONICS SEV", 
            "DYM VIETNAM CO., LTD"
        ],
        'CONSIGNEE': [
            "NIPPON EXPRESS KOREA", 
            "KUEHNE + NAGEL KOREA", 
            "SAMSUNG HQ SEOUL", 
            "CJ LOGISTICS SEOUL"
        ],
        'W/T': [165.5, 85.0, 520.0, 12.4],
        'R.W/T': [158.0, 92.5, 537.2, 15.0],  # WEIGHT tính toán lấy từ đây
        'V.W': [140.0, 92.5, 490.0, 10.0]
    }
    
    df_raw = pd.DataFrame(data_hinh_1)
    
    # Bước 1: Xử lý quy chuẩn và phân tích tự động
    print("\\nBước 1: Tự động phân tích & mapping dữ liệu tuyến đường (ROUTE) và Trọng lượng...")
    df_processed = process_manifest(df_raw)
    print(df_processed[['MAWB NO', 'ROUTE', 'WEIGHT', 'HANDLING_CHARGE', 'WAREHOUSE_CHARGE']])
    
    # Bước 2: Thiết lập cấu hình đầu vào từ người dùng
    shipper_target = "DYM VIETNAM CO., LTD"
    ty_gia_exchange = 26160
    
    # Từ điển đơn giá USD/kg nhập tay cho từng dòng hàng hóa (sử dụng HAWB làm lock-key)
    don_gia_nhap_tay = {
        "HW-WH-45": 3.40,  # Đơn giá dòng 1
        "HW-WH-25": 3.40,  # Đơn giá dòng 2
        "HW-WH-12": 4.10,  # Đơn giá dòng 4
    }
    
    # Bước 3: Xuất báo cáo hoàn chỉnh ra file Excel giống Hình 2
    output_filename = "Bang_Ke_Hoa_Don_Hinh2_Python_Output.xlsx"
    print(f"\\nBước 2: Tiến hành kết xuất báo cáo Excel cho '{shipper_target}'...")
    export_shipper_report(
        df_processed=df_processed,
        shipper_name=shipper_target,
        unit_prices_dict=don_gia_nhap_tay,
        exchange_rate=ty_gia_exchange,
        output_path=output_filename,
        month=6,
        year=2026
    )
    print("--- Hoàn tất! Bạn có thể kiểm tra tệp Excel đầu ra ---")
`;

  const handleCopy = () => {
    navigator.clipboard.writeText(pythonCode);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="bg-slate-900 text-slate-100 rounded-2xl p-6 shadow-xl border border-slate-800">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-6 pb-6 border-b border-slate-800">
        <div className="flex items-center gap-3">
          <div className="p-2.5 bg-emerald-500/10 rounded-xl text-emerald-400">
            <FileCode2 className="w-6 h-6" />
          </div>
          <div>
            <h3 className="font-semibold text-lg text-white">Mã nguồn Python Hoàn Chỉnh</h3>
            <p className="text-xs text-slate-400">Sử dụng Pandas để xử lý dữ liệu và OpenPyXL để kết xuất báo cáo Excel định dạng tiêu chuẩn</p>
          </div>
        </div>
        
        <div className="flex gap-2 w-full md:w-auto">
          <button
            onClick={handleCopy}
            className="flex-1 md:flex-none flex items-center justify-center gap-2 px-3.5 py-1.5 bg-slate-800 hover:bg-slate-700 active:bg-slate-600 text-xs font-medium text-slate-200 hover:text-white rounded-lg transition-colors border border-slate-700"
          >
            {copied ? (
              <>
                <Check className="w-3.5 h-3.5 text-green-400" />
                <span className="text-green-400">Đã sao chép!</span>
              </>
            ) : (
              <>
                <Copy className="w-3.5 h-3.5" />
                <span>Sao chép Python Code</span>
              </>
            )}
          </button>
        </div>
      </div>

      <div className="bg-sky-500/5 text-sky-450 border border-sky-500/10 rounded-xl p-4 mb-6 text-xs flex gap-3 leading-relaxed">
        <Info className="w-5 h-5 text-sky-400 shrink-0 mt-0.5" />
        <div>
          <span className="font-semibold block text-slate-200 mb-1">Hướng dẫn thiết lập môi trường:</span>
          Cài đặt các thư viện cần thiết bằng câu lệnh: <code className="bg-slate-950 px-1.5 py-0.5 rounded text-rose-400 font-mono">pip install pandas openpyxl</code> trước khi chạy script.
          Script này tự động sinh file dữ liệu mô phỏng, phân tích tự động tuyến bay và xây dựng trực tiếp file Excel được styling đầy đủ màu nền Mint của hàng tiêu đề, màu vàng ấm của cột Tổng cùng định dạng VND chuyên nghiệp cho dòng cuối cùng.
        </div>
      </div>

      <h4 className="text-sm font-semibold text-slate-200 mb-2 flex items-center gap-2">python_manifest_billing.py</h4>
      <div className="bg-slate-950 rounded-xl p-4 overflow-x-auto border border-slate-800 text-xs leading-relaxed max-h-[450px]">
        <pre className="font-mono text-amber-400">
          <code>{pythonCode}</code>
        </pre>
      </div>
    </div>
  );
}
