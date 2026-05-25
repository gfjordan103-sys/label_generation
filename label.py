import streamlit as st
import openpyxl
from openpyxl.styles import Font, Border, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import io

# ==========================================
# 核心邏輯：20欄一頁完美Cycle，欄寬、列高、格式完全同步樣版
# ==========================================
def generate_exact_printing_cycle_labels(detail_wb, template_wb):
    ws_detail = detail_wb.worksheets[0]
    ws_template = template_wb.worksheets[0]
    
    # 1. 解析出貨明細：加總各板號箱數與對應訂單號碼
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

    # 2. 建立全新的列印活頁簿
    output_wb = openpyxl.Workbook()
    ws_output = output_wb.active
    ws_output.title = "標籤自動排版列印表"
    
    # 【高精準修正：列高同步】一比一複製 1~34 列的精確列高，確保垂直不跑版
    for r_idx in range(1, 35):
        if r_idx in ws_template.row_dimensions and ws_template.row_dimensions[r_idx].height is not None:
            ws_output.row_dimensions[r_idx].height = ws_template.row_dimensions[r_idx].height

    # 【高精準修正：欄寬緩存】一比一讀取並緩存樣版 A(1) 到 T(20) 的原始欄寬
    template_widths = {}
    for c_idx in range(1, 21):
        col_letter = get_column_letter(c_idx)
        if col_letter in ws_template.column_dimensions and ws_template.column_dimensions[col_letter].width is not None:
            template_widths[c_idx] = ws_template.column_dimensions[col_letter].width
        else:
            template_widths[c_idx] = 10.0

    # 複製特定格子網格與格式的輔住函式
    def apply_style_from_template(target_cell, r, t_c):
        src_cell = ws_template.cell(row=r, column=t_c)
        if src_cell.has_style:
            if src_cell.font: target_cell.font = Font(**src_cell.font.__dict__)
            if src_cell.border: target_cell.border = Border(**src_cell.border.__dict__)
            if src_cell.fill: target_cell.fill = PatternFill(**src_cell.fill.__dict__)
            if src_cell.alignment: target_cell.alignment = Alignment(**src_cell.alignment.__dict__)
            target_cell.number_format = src_cell.number_format

    # 3. 橫向 Cycle 排版引擎
    # 每一頁 20 欄中，7組標籤的相對起始欄位索引（1=A, 4=D, 7=G, 10=J, 13=M, 16=P, 19=S）
    label_group_offsets = [1, 4, 7, 10, 13, 16, 19]
    current_group_global_index = 0  
    MAX_LABELS_PER_COL = 32  # 限制：第 1 列留白，第 2 列板號，下方最多放 32 列標籤（總高限34列）

    for board_name in board_data:
        info = board_dict[board_name]
        boxes_count = info['total_boxes']
        order_no = info['order_no']
        
        if boxes_count <= 0:
            continue
            
        remaining_boxes = boxes_count
        
        while remaining_boxes > 0:
            current_col_boxes = min(remaining_boxes, MAX_LABELS_PER_COL)
            
            # 計算當前組別在無限橫向延伸下的頁數與絕對欄位位置
            page_number = current_group_global_index // 7   
            group_in_page = current_group_global_index % 7  
            start_col_idx = (page_number * 20) + label_group_offsets[group_in_page]
            
            # A. 寫入第 2 列：板號名稱（主內容欄與並排欄都套用樣式）
            for c_offset in range(2):
                target_c = start_col_idx + c_offset
                tpl_c = label_group_offsets[group_in_page] + c_offset
                
                h_cell = ws_output.cell(row=2, column=target_c)
                apply_style_from_template(h_cell, 2, tpl_c)
                if c_offset == 0:
                    h_cell.value = board_name

            # B. 寫入第 3 到第 34 列：訂單號碼標籤
            for i in range(1, current_col_boxes + 1):
                target_row = 2 + i  
                for c_offset in range(2):
                    target_c = start_col_idx + c_offset
                    tpl_c = label_group_offsets[group_in_page] + c_offset
                    
                    cell = ws_output.cell(row=target_row, column=target_c)
                    apply_style_from_template(cell, target_row, tpl_c)
                    if c_offset == 0:
                        cell.value = order_no
            
            # C. 防呆補線：如果箱數沒裝滿 32 列，剩餘的空白格子也要補齊網格線，維持美觀
            for i in range(current_col_boxes + 1, MAX_LABELS_PER_COL + 1):
                target_row = 2 + i
                for c_offset in range(2):
                    target_c = start_col_idx + c_offset
                    tpl_c = label_group_offsets[group_in_page] + c_offset
                    cell = ws_output.cell(row=target_row, column=target_c)
                    apply_style_from_template(cell, target_row, tpl_c)

            remaining_boxes -= current_col_boxes
            current_group_global_index += 1

    # 【20欄 Cycle 欄寬強制造型外推】確保任何新增的寬窄間隔（如 C, F, I 以及跨頁後的欄位）都完美同步
    total_used_cols = ((current_group_global_index - 1) // 7 + 1) * 20
    for c_idx in range(1, total_used_cols + 1):
        c_let = get_column_letter(c_idx)
        tpl_idx = ((c_idx - 1) % 20) + 1
        ws_output.column_dimensions[c_let].width = template_widths[tpl_idx]

    return output_wb

# ==========================================
# Streamlit 純網頁 UI 介面
# ==========================================
st.set_page_config(page_title="印刷標籤自動排版系統", page_icon="🖨️")
st.title("🖨️ 印刷標籤自動排版系統")

file_detail = st.file_uploader("1. 請上傳【出貨明細】Excel 檔案 (.xlsx)", type=["xlsx"], key="p_detail")
file_template = st.file_uploader("2. 請上傳【訂單號碼】標籤樣版 Excel 檔案 (.xlsx)", type=["xlsx"], key="p_template")

if file_detail and file_template:
    st.success("📊 雙檔案載入成功！")
    if st.button("🚀 執行印刷排版", type="primary"):
        with st.spinner("正在進行精密橫向分頁對位與欄寬列高同步..."):
            try:
                wb_d = openpyxl.load_workbook(file_detail)
                wb_t = openpyxl.load_workbook(file_template)
                
                # 【修正：呼叫正確的名字】
                result_wb = generate_exact_printing_cycle_labels(wb_d, wb_t)
                
                excel_buffer = io.BytesIO()
                result_wb.save(excel_buffer)
                excel_buffer.seek(0)
                
                st.success("🎉 排版成功！欄寬與列高已完美實現 Cycle 複製，錯誤已完全修復。")
                st.download_button(
                    label="📥 下載印刷對位標籤結果",
                    data=excel_buffer,
                    file_name="印刷標籤結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ 錯誤：{e}")
