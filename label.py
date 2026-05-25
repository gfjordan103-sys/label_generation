import streamlit as st
import openpyxl
from openpyxl.styles import Font, Border, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import io

# ==========================================
# 核心邏輯：依樣版完美複製欄寬、格式，並執行限高34列向右排版
# ==========================================
def generate_exact_matrix_labels(detail_wb, template_wb):
    ws_detail = detail_wb.worksheets[0]
    ws_template = template_wb.worksheets[0]
    
    # 1. 步驟一：解析出貨明細，記錄各板號的總箱數與訂單號碼
    board_data = []
    board_dict = {}
    current_board = None
    
    for row in range(3, ws_detail.max_row + 1):
        board_val = ws_detail.cell(row=row, column=1).value  # A欄：板數
        box_val = ws_detail.cell(row=row, column=7).value    # G欄：箱數
        order_val = ws_detail.cell(row=row, column=8).value  # H欄：訂單號碼
        
        if board_val and str(board_val).strip():
            current_board = str(board_val).strip()
            if current_board not in board_dict:
                board_dict[current_board] = {'total_boxes': 0, 'order_no': ''}
                board_data.append(current_board)
        
        if current_board and box_val is not None:
            try:
                board_dict[current_board]['total_boxes'] += int(float(box_val))
            except ValueError:
                pass
            if order_val and not board_dict[current_board]['order_no']:
                board_dict[current_board]['order_no'] = str(order_val).split('.')[0]

    # 2. 步驟二：建立全新的空白活頁簿
    output_wb = openpyxl.Workbook()
    ws_output = output_wb.active
    ws_output.title = "標籤列印表"
    
    # 【核心修正】直接一比一複製樣版中所有欄位的寬度，C、F、I欄不對稱寬度完美保留！
    for col_idx in range(1, ws_template.max_column + 1):
        col_letter = get_column_letter(col_idx)
        if col_letter in ws_template.column_dimensions:
            ws_output.column_dimensions[col_letter].width = ws_template.column_dimensions[col_letter].width
        else:
            # 防呆：如果樣版後面沒有定義，複製前幾欄的寬度
            ref_col = get_column_letter(((col_idx - 1) % 3) + 1)
            if ref_col in ws_template.column_dimensions:
                ws_output.column_dimensions[col_letter].width = ws_template.column_dimensions[ref_col].width

    # 抓取樣版基礎樣式（以第一組 A2、A3 為基準，完美保留網格與格式）
    ref_hdr_cell = ws_template.cell(row=2, column=1) # 樣版 A2 (板號)
    ref_lbl_cell = ws_template.cell(row=3, column=1) # 樣版 A3 (標籤)
    
    hdr_font = Font(**ref_hdr_cell.font.__dict__) if ref_hdr_cell.font else Font(name='Arial', size=12, bold=True)
    lbl_font = Font(**ref_lbl_cell.font.__dict__) if ref_lbl_cell.font else Font(name='Arial', size=11)
    
    ref_border = Border(**ref_lbl_cell.border.__dict__) if ref_lbl_cell.border else Border()
    ref_fill = PatternFill(**ref_lbl_cell.fill.__dict__) if ref_lbl_cell.fill else PatternFill()
    ref_alignment = Alignment(**ref_lbl_cell.alignment.__dict__) if ref_lbl_cell.alignment else Alignment(horizontal='center', vertical='center')

    # 3. 步驟三：由左往右進行矩陣排版
    current_start_col = 1  # 第一組從 A 欄(第1欄)開始，佔用 AB 欄
    MAX_LABELS_PER_COL = 32  # 限制：第1列標頭+第2列板號，下方最多放 32 列訂單號碼（剛好到第34列）
    total_labels_written = 0

    for board_name in board_data:
        info = board_dict[board_name]
        boxes_count = info['total_boxes']
        order_no = info['order_no']
        
        if boxes_count <= 0:
            continue
            
        remaining_boxes = boxes_count
        
        while remaining_boxes > 0:
            # 這一欄要填入的標籤數（最多 32 箱）
            current_col_boxes = min(remaining_boxes, MAX_LABELS_PER_COL)
            
            # A. 寫入第 2 列：板號名稱（主欄與右邊併排欄都套用樣式與網格邊框）
            for c_offset in range(2):
                h_cell = ws_output.cell(row=2, column=current_start_col + c_offset)
                h_cell.font = hdr_font
                h_cell.border = ref_border
                h_cell.fill = ref_fill
                h_cell.alignment = ref_alignment
                if c_offset == 0:
                    h_cell.value = board_name # 僅在左側主欄寫入板號名稱
            
            # B. 從第 3 列開始往下填寫訂單號碼標籤，精準卡在第 34 列上限
            for i in range(1, current_col_boxes + 1):
                target_row = 2 + i  # 3, 4, ..., 34
                
                # 填寫主內容欄（如 A 欄）與並排格子欄（如 B 欄）
                for c_offset in range(2):
                    cell = ws_output.cell(row=target_row, column=current_start_col + c_offset)
                    cell.font = lbl_font
                    cell.border = ref_border
                    cell.fill = ref_fill
                    cell.alignment = ref_alignment
                    if c_offset == 0:
                        cell.value = order_no # 僅在左側主欄寫入對應的訂單號碼
                
                total_labels_written += 1
                
            remaining_boxes -= current_col_boxes
            # 換到右邊下一個標籤區塊（AB欄(1,2) -> 跳過C(3) -> DE欄(4,5) -> 跳過F(6) -> GH欄(7,8)...）
            current_start_col += 3 
            
    return output_wb, total_labels_written

# ==========================================
# Streamlit 純網頁 UI 介面
# ==========================================
st.set_page_config(page_title="標籤排版系統", page_icon="🏷️", layout="centered")

st.title("🏷️ 標籤排版系統")
st.write("【排版規範：AB一組、DE一組，限高 34 列】。一比一複製樣版所有欄寬（包含不對稱之 C, F, I 欄），格式完全不跑偏。")

# 網頁 UI 檔案上傳區
file_detail = st.file_uploader("1. 請上傳【出貨明細】Excel 檔案 (.xlsx)", type=["xlsx"], key="web_detail")
file_template = st.file_uploader("2. 請上傳【訂單號碼】標籤樣版 Excel 檔案 (.xlsx)", type=["xlsx"], key="web_template")

if file_detail and file_template:
    st.success("📊 兩份實體檔案已成功載入網頁 UI 緩衝區！")
    
    if st.button("🚀 開始實體規格排版生成", type="primary"):
        with st.spinner("正在進行精確對位並一比一轉移欄寬與格式..."):
            try:
                wb_d = openpyxl.load_workbook(file_detail)
                wb_t = openpyxl.load_workbook(file_template)
                
                # 執行全新對位與欄寬完美同步邏輯
                result_wb, total_count = generate_exact_matrix_labels(wb_d, wb_t)
                
                # 轉為前端下載流
                excel_buffer = io.BytesIO()
                result_wb.save(excel_buffer)
                excel_buffer.seek(0)
                
                st.success(f"🎉 UI Demo 成功！已自動填入 {total_count} 列標籤，並已依樣版格式一比一完美同步 C, F, I 等欄寬！")
                
                # 下載按鈕
                st.download_button(
                    label="📥 下載精確對位標籤 Excel",
                    data=excel_buffer,
                    file_name="實體規格對位標籤結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ 處理過程中發生錯誤，請確認 Excel 內容：{e}")