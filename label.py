import streamlit as st
import openpyxl
from openpyxl.styles import Font, Border, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import io

# ==========================================
# 核心邏輯：20欄一頁完美循環，精確對位不走鐘
# ==========================================
def generate_printing_page_labels(detail_wb, template_wb):
    ws_detail = detail_wb.worksheets[0]
    ws_template = template_wb.worksheets[0]
    
    # 1. 解析出貨明細
    board_data = []
    board_dict = {}
    current_board = None
    
    for row in range(3, ws_detail.max_row + 1):
        board_val = ws_detail.cell(row=row, column=1).value  # A欄
        box_val = ws_detail.cell(row=row, column=7).value    # G欄
        order_val = ws_detail.cell(row=row, column=8).value  # H欄
        
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

    # 2. 建立全新的列印活頁簿
    output_wb = openpyxl.Workbook()
    ws_output = output_wb.active
    ws_output.title = "標籤自動排版列印表"
    
    # 一比一紀錄樣版 A(1) 到 T(20) 的【神聖不可侵犯欄寬】
    template_widths = {}
    for c_idx in range(1, 21):
        col_letter = get_column_letter(c_idx)
        # 確保抓到樣版的真實寬度，抓不到就給 10
        if col_letter in ws_template.column_dimensions and ws_template.column_dimensions[col_letter].width is not None:
            template_widths[c_idx] = ws_template.column_dimensions[col_letter].width
        else:
            template_widths[c_idx] = 10.0

    # 樣版基礎元件樣式參考（直接抓取樣版 A2, A3 的框線與字體）
    ref_hdr_cell = ws_template.cell(row=2, column=1)
    ref_lbl_cell = ws_template.cell(row=3, column=1)
    
    hdr_font = Font(**ref_hdr_cell.font.__dict__) if ref_hdr_cell.font else Font(name='Arial', size=12, bold=True)
    lbl_font = Font(**ref_lbl_cell.font.__dict__) if ref_lbl_cell.font else Font(name='Arial', size=11)
    ref_border = Border(**ref_lbl_cell.border.__dict__) if ref_lbl_cell.border else Border()
    ref_fill = PatternFill(**ref_lbl_cell.fill.__dict__) if ref_lbl_cell.fill else PatternFill()
    ref_alignment = Alignment(**ref_lbl_cell.alignment.__dict__) if ref_lbl_cell.alignment else Alignment(horizontal='center', vertical='center')

    # 每頁 20 欄中，7組標籤的「相對起始欄位索引」（1=A, 4=D, 7=G, 10=J, 13=M, 16=P, 19=S）
    label_group_offsets = [1, 4, 7, 10, 13, 16, 19]
    current_group_global_index = 0  
    MAX_LABELS_PER_COL = 32 # 扣除1、2列，最多填 32 列標籤

    # 3. 排版引擎
    for board_name in board_data:
        info = board_dict[board_name]
        boxes_count = info['total_boxes']
        order_no = info['order_no']
        
        if boxes_count <= 0:
            continue
            
        remaining_boxes = boxes_count
        
        while remaining_boxes > 0:
            current_col_boxes = min(remaining_boxes, MAX_LABELS_PER_COL)
            
            # 計算這組標籤應該在第幾頁、該頁的第幾組
            page_number = current_group_global_index // 7   
            group_in_page = current_group_global_index % 7  
            
            # 算出絕對起點欄位
            start_col_idx = (page_number * 20) + label_group_offsets[group_in_page]
            
            # A. 寫入第 2 列板號
            for c_offset in range(2):
                h_cell = ws_output.cell(row=2, column=start_col_idx + c_offset)
                h_cell.font = hdr_font
                h_cell.border = ref_border
                h_cell.fill = ref_fill
                h_cell.alignment = ref_alignment
                if c_offset == 0:
                    h_cell.value = board_name

            # B. 寫入第 3~34 列訂單號碼
            for i in range(1, current_col_boxes + 1):
                target_row = 2 + i  
                for c_offset in range(2):
                    cell = ws_output.cell(row=target_row, column=start_col_idx + c_offset)
                    cell.font = lbl_font
                    cell.border = ref_border
                    cell.fill = ref_fill
                    cell.alignment = ref_alignment
                    if c_offset == 0:
                        cell.value = order_no
            
            remaining_boxes -= current_col_boxes
            current_group_global_index += 1

    # 【核心修正】等所有標籤畫完，用最強硬的「20欄模數循環」強行重設整張表所有欄位的寬度！
    total_used_cols = ((current_group_global_index - 1) // 7 + 1) * 20
    for c_idx in range(1, total_used_cols + 1):
        c_let = get_column_letter(c_idx)
        # 精準對應：不管是第幾頁，第 c_idx 欄的寬度永遠等於對應第一頁 (1~20) 的寬度
        tpl_idx = ((c_idx - 1) % 20) + 1
        ws_output.column_dimensions[c_let].width = template_widths[tpl_idx]

    return output_wb

# ==========================================
# Streamlit UI
# ==========================================
st.set_page_config(page_title="印刷標籤自動排版系統", page_icon="🖨️")
st.title("🖨️ 印刷標籤自動排版系統 (欄寬修正版)")
st.write("已修正間隔欄(C, F, I)縮小跑版問題，完美保留原始範例 Excel 的所有寬度。")

file_detail = st.file_uploader("1. 上傳【出貨明細】(.xlsx)", type=["xlsx"], key="p_detail")
file_template = st.file_uploader("2. 上傳【訂單號碼】標籤樣版 (.xlsx)", type=["xlsx"], key="p_template")

if file_detail and file_template:
    st.success("📊 檔案已成功載入！")
    if st.button("🚀 執行 20 欄 Cycle 印刷排版", type="primary"):
        with st.spinner("正在強制同步 C, F, I 等間隔欄寬..."):
            try:
                wb_d = openpyxl.load_workbook(file_detail)
                wb_t = openpyxl.load_workbook(file_template)
                result_wb = generate_printing_page_labels(wb_d, wb_t)
                
                excel_buffer = io.BytesIO()
                result_wb.save(excel_buffer)
                excel_buffer.seek(0)
                
                st.success("🎉 排版修正成功！間隔欄已恢復特製窄度，請點下方按鈕下載。")
                st.download_button(
                    label="📥 下載修正後的標籤 Excel",
                    data=excel_buffer,
                    file_name="修正後_印刷對位標籤結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ 錯誤：{e}")
