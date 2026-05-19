import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- CẤU HÌNH GIAO DIỆN WEB ---
st.set_page_config(page_title="Hệ thống Xử lý & Phân tích Dữ liệu Phát hành", layout="wide")
st.title("📊 HỆ THỐNG XỬ LÝ DỮ LIỆU PHÁT HÀNH SÁCH TỰ ĐỘNG")
st.write("Tích hợp quy trình: Làm sạch dữ liệu dòng thừa ➡️ Phân tích, gom nhóm mã hàng ➡️ Tách hóa đơn tự động.")

# --- HÀM TRANG TRÍ EXCEL CHUYÊN NGHIỆP ---
def trang_tri_sheet(worksheet, tieude_color, has_vat_summary=False, total_row_type="standard"):
    font_tieude = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fill_tieude = PatternFill(start_color=tieude_color, end_color=tieude_color, fill_type="solid")
    font_noidung = Font(name="Arial", size=10)
    font_tongket = Font(name="Arial", size=10, bold=True)
    
    vien_mong = Side(border_style="thin", color="D9D9D9")
    border_o = Border(left=vien_mong, right=vien_mong, top=vien_mong, bottom=vien_mong)
    
    # 1. Định dạng dòng tiêu đề
    worksheet.row_dimensions[1].height = 25
    for col_idx in range(1, worksheet.max_column + 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.font = font_tieude
        cell.fill = fill_tieude
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_o

    # 2. Định dạng nội dung và số phân ngàn
    for row_idx in range(2, worksheet.max_row + 1):
        worksheet.row_dimensions[row_idx].height = 19
        
        # Kiểm tra nếu là các dòng tổng kết cuối bảng
        is_total_row = False
        if total_row_type == "grand" and row_idx == worksheet.max_row:
            is_total_row = True
        elif total_row_type == "standard" and has_vat_summary and row_idx >= worksheet.max_row - 2:
            is_total_row = True
        elif total_row_type == "standard" and not has_vat_summary and row_idx == worksheet.max_row:
            is_total_row = True

        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.border = border_o
            
            if is_total_row:
                cell.font = font_tongket
                if total_row_type == "grand":
                    cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Màu vàng nhạt cho Doanh thu thuần
            else:
                cell.font = font_noidung

            # Định dạng số phân ngàn cho các cột tiền/số lượng
            if col_idx in [1, 3]: # STT, Mã số
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col_idx == 3: cell.number_format = '@'
            elif col_idx == 2: # Tên sách
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_idx == 4: # ĐVT
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else: # Các cột số tính toán
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0"

    # 3. Tự động giãn cột
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 11)

# --- KHỞI CHẠY QUY TRÌNH XỬ LÝ ---
uploaded_file = st.file_uploader("Kéo thả file Excel thô của hệ thống vào đây:", type=["xlsx"])

if uploaded_file is not None:
    st.success("Đã nhận file thô thành công!")
    
    if st.button("🚀 BẮT ĐẦU XỬ LÝ VÀ PHÂN TÍCH MASTER PROCESS"):
        try:
            with st.spinner("Hệ thống đang thực hiện bộ lọc sạch và phân tích tài chính ngầm..."):
                excel_file = pd.ExcelFile(uploaded_file)
                all_sheets_data = []
                
                # --- BƯỚC 1: LÀM SẠCH VÀ GỘP DỮ LIỆU THÔ ---
                for sheet_name in excel_file.sheet_names:
                    df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=8, header=None)
                    if df_sheet.empty: continue
                    
                    df_sheet.columns = ['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']
                    
                    for col in ['Tên sách', 'Mã số', 'ĐVT']:
                        df_sheet[col] = df_sheet[col].astype(str).str.strip()
                    
                    df_sheet = df_sheet[df_sheet['Mã số'].notna() & (df_sheet['Mã số'] != '') & (df_sheet['Mã số'] != 'nan') & (df_sheet['Mã số'] != 'Mã số')]
                    tu_khoa_xoa = ['Tổng cộng:', 'Thuế VAT:', 'Thành tiền:', 'Tổng cộng', 'Thuế VAT', 'Phụ trách cung tiêu', 'Người giao hàng', 'Thủ Kho', 'nan', 'STT', 'Tên sách']
                    df_sheet = df_sheet[~df_sheet['Tên sách'].isin(tu_khoa_xoa)]
                    df_sheet = df_sheet[pd.to_numeric(df_sheet['Số lượng'], errors='coerce').notna()]
                    
                    all_sheets_data.append(df_sheet)
                
                if not all_sheets_data:
                    st.error("Không trích xuất được dữ liệu hợp lệ từ file.")
                    st.stop()
                    
                df_raw = pd.concat(all_sheets_data, ignore_index=True)
                for col in ['Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']:
                    df_raw[col] = pd.to_numeric(df_raw[col])
                df_raw['STT'] = range(1, len(df_raw) + 1)
                
                # --- BƯỚC 2: PHÂN TÍCH MASTER PROCESS (Dựa trên file .bas) ---
                df_duong = df_raw[df_raw['Số lượng'] >= 0].copy()
                df_am = df_raw[df_raw['Số lượng'] < 0].copy()
                
                # 2.1 Tổng hợp mã hàng bán (Dương - Không VAT)
                df_tonghop_ban = df_duong.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa'], as_index=False).agg({'Số lượng': 'sum'})
                df_tonghop_ban['CK'] = 0
                df_tonghop_ban['Đơn giá'] = df_tonghop_ban['Giá bìa']
                df_tonghop_ban['Thành tiền'] = df_tonghop_ban['Số lượng'] * df_tonghop_ban['Đơn giá']
                df_tonghop_ban = df_tonghop_ban[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                df_tonghop_ban.insert(0, 'STT', range(1, len(df_tonghop_ban) + 1))
                
                # 2.2 Tổng hợp mã hàng trả (Biến đổi Âm thành Dương - Không VAT)
                df_tonghop_tra = df_am.copy()
                df_tonghop_tra['Số lượng'] = df_tonghop_tra['Số lượng'].abs()
                df_tonghop_tra = df_tonghop_tra.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa'], as_index=False).agg({'Số lượng': 'sum'})
                df_tonghop_tra['CK'] = 0
                df_tonghop_tra['Đơn giá'] = df_tonghop_tra['Giá bìa']
                df_tonghop_tra['Thành tiền'] = df_tonghop_tra['Số lượng'] * df_tonghop_tra['Đơn giá']
                df_tonghop_tra = df_tonghop_tra[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                df_tonghop_tra.insert(0, 'STT', range(1, len(df_tonghop_tra) + 1))
                
                # 2.3 Chi tiết âm (Thực tế giữ nguyên dấu âm để tính toán khấu trừ)
                df_chitiet_am = df_am.copy()
                df_chitiet_am['Thành tiền'] = df_chitiet_am['Số lượng'] * df_chitiet_am['Đơn giá']
                df_chitiet_am['STT'] = range(1, len(df_chitiet_am) + 1)
                
                # 2.4 Toàn bộ hàng dương xuất ra
                df_tonghop_banra = df_duong.copy()
                df_tonghop_banra['STT'] = range(1, len(df_tonghop_banra) + 1)
                
                # --- QUY TRÌNH ĐẨY DỮ LIỆU RA CẤU TRÚC EXCEL NHIỀU SHEET ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    
                    # Tạo cấu trúc lưu trữ dữ liệu báo cáo tài chính tổng kết chung
                    grand_summary_data = []
                    
                    # Hàm phụ để ghi dữ liệu và tính dòng tổng kết
                    def ghi_sheet_kem_tong(df, name, color, has_vat=False):
                        df.to_excel(writer, index=False, sheet_name=name)
                        ws = writer.sheets[name]
                        r_idx = len(df) + 2
                        
                        sum_sl = df['Số lượng'].sum()
                        sum_tt = df['Thành tiền'].sum()
                        
                        ws.cell(row=r_idx, column=6).value = "Tổng cộng:"
                        ws.cell(row=r_idx, column=7).value = sum_sl
                        ws.cell(row=r_idx, column=9).value = sum_tt
                        
                        vat, sau_thue = 0.0, 0.0
                        if has_vat:
                            vat = sum_tt * 0.05
                            sau_thue = sum_tt * 1.05
                            ws.cell(row=r_idx+1, column=8).value = "VAT 5%:"
                            ws.cell(row=r_idx+1, column=9).value = vat
                            ws.cell(row=r_idx+2, column=8).value = "Sau Thuế:"
                            ws.cell(row=r_idx+2, column=9).value = sau_thue
                            
                        trang_tri_sheet(ws, color, has_vat_summary=has_vat)
                        return len(df), sum_sl, sum_tt, vat, sau_thue

                    # Ghi Sheet 1: Tổng hợp mã hàng bán
                    if not df_tonghop_ban.empty:
                        ln, sl, tt, _, _ = ghi_sheet_kem_tong(df_tonghop_ban, "Tong_Hop_Ma_Hang_Ban", "003366")
                        grand_summary_data.append(["Tổng Hợp Mã Hàng Bán", "Dương", ln, sl, tt, 0, tt, "Giá Bìa (Không VAT)"])
                        
                    # Ghi Sheet 2: Tổng hợp hàng trả
                    if not df_tonghop_tra.empty:
                        ln, sl, tt, _, _ = ghi_sheet_kem_tong(df_tonghop_tra, "Tong_Hop_Hang_Tra", "FF6600")
                        grand_summary_data.append(["Tổng Hợp Hàng Trả", "Dương", ln, sl, tt, 0, tt, "Giá Bìa (Không VAT)"])
                        
                    # Ghi Sheet 3: Chi tiết âm
                    if not df_chitiet_am.empty:
                        ln, sl, tt, vt, st = ghi_sheet_kem_tong(df_chitiet_am, "Chi_Tiet_Am", "C00000", has_vat=True)
                        grand_summary_data.append(["Chi Tiết Hàng Trả", "Âm", ln, sl, tt, vt, st, "Thực tế sau CK"])
                        
                    # Ghi Sheet 4: Tổng hợp bán ra (Toàn bộ dữ liệu dương)
                    if not df_tonghop_banra.empty:
                        ln, sl, tt, vt, st = ghi_sheet_kem_tong(df_tonghop_banra, "Tong_Hop_Ban_Ra", "00B050", has_vat=True)
                        grand_summary_data.append(["Tổng Hợp Bán Ra", "Dương", ln, sl, tt, vt, st, "Toàn bộ dữ liệu dương"])

                    # TÁCH HOÁ ĐƠN: Cứ 1000 dòng hàng dương cắt ra thành 1 sheet HD
                    batch_size = 1000
                    sheet_index = 1
                    for i in range(0, len(df_tonghop_banra), batch_size):
                        df_batch = df_tonghop_banra.iloc[i:i+batch_size].copy()
                        df_batch['STT'] = range(1, len(df_batch) + 1)
                        hd_name = f"HD {sheet_index}"
                        ln, sl, tt, vt, st = ghi_sheet_kem_tong(df_batch, hd_name, "0070C0", has_vat=True)
                        grand_summary_data.append([f"Hóa Đơn {sheet_index}", "Dương", ln, sl, tt, vt, st, "Tách lẻ 1000 dòng"])
                        sheet_index += 1

                    # TẠO SHEET: TỔNG KẾT CHUNG (Đặt ở đầu bảng tính)
                    df_grand = pd.DataFrame(grand_summary_data, columns=["Tên Hạng mục", "Loại", "Số dòng", "Số lượng", "Trước Thuế", "VAT (5%)", "Sau Thuế", "Ghi chú"])
                    df_grand.to_excel(writer, index=False, sheet_name="Tong_Ket_Chung")
                    ws_grand = writer.sheets["Tong_Ket_Chung"]
                    
                    # Tính doanh thu thuần thực tế = Tổng Dương + Tổng Âm (Do hàng âm giữ nguyên dấu âm nên dùng phép cộng)
                    thuong_sl = df_raw['Số lượng'].sum()
                    thuong_tt = df_raw['Thành tiền'].sum()
                    thuong_vat = thuong_tt * 0.05
                    thuong_st = thuong_tt * 1.05
                    
                    r_grand = len(df_grand) + 2
                    ws_grand.cell(row=r_grand, column=1).value = "DOANH THU THỰC TẾ (BAN - TRA)"
                    ws_grand.cell(row=r_grand, column=4).value = thuong_sl
                    ws_grand.cell(row=r_grand, column=5).value = thuong_tt
                    ws_grand.cell(row=r_grand, column=6).value = thuong_vat
                    ws_grand.cell(row=r_grand, column=7).value = thuong_st
                    
                    trang_tri_sheet(ws_grand, "000000", total_row_type="grand")
                    
                    # Đẩy sheet tổng kết lên vị trí đầu tiên
                    wb = writer.book
                    order = [wb.sheetnames[-1]] + wb.sheetnames[:-1]
                    wb._sheets = [wb._sheets[wb.sheetnames.index(name)] for name in order]

                processed_data = output.getvalue()
                st.success("🎉 PHÂN TÍCH THÀNH CÔNG MASTER PROCESS!")
                
                # --- HIỂN THỊ TRỰC QUAN BẰNG CÁC TAB TRÊN WEB ---
                tab1, tab2, tab3, tab4 = st.tabs(["📋 Tổng Kết Chung", "📈 Tổng Hợp Mã Bán", "📉 Tổng Hợp Mã Trả", "📦 Chi Tiết Hàng Âm"])
                with tab1:
                    st.dataframe(df_grand, use_container_width=True)
                    st.metric(label="Tổng Doanh Thu Thực Tế (Sau Thuế)", value=f"{thuong_st:,.0f} đ")
                with tab2:
                    st.dataframe(df_tonghop_ban.head(50), use_container_width=True)
                with tab3:
                    st.dataframe(df_tonghop_tra.head(50), use_container_width=True)
                with tab4:
                    st.dataframe(df_chitiet_am.head(50), use_container_width=True)

                # Nút tải file
                st.download_button(
                    label="📥 TẢI FILE EXCEL MASTER REPORT (FULL SHEETS)",
                    data=processed_data,
                    file_name=f"Master_Report_Sach_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Lỗi hệ thống trong quy trình xử lý: {str(e)}")
