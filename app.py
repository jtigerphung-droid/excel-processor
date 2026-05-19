import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- CẤU HÌNH GIAO DIỆN WEB ---
st.set_page_config(page_title="Hệ thống Xử lý & Phân tích Dữ liệu Phát hành", layout="wide")
st.title("📊 HỆ THỐNG XỬ LÝ DỮ LIỆU PHÁT HÀNH SÁCH TỰ ĐỘNG")
st.write("Phiên bản Thuật toán Định vị Cột Động: Chống lỗi lệch cấu trúc file và thiếu cột hệ thống.")

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
    max_r = worksheet.max_row
    for row_idx in range(2, max_r + 1):
        worksheet.row_dimensions[row_idx].height = 19
        
        is_total_row = False
        if total_row_type == "grand" and row_idx == max_r:
            is_total_row = True
        elif total_row_type == "standard" and has_vat_summary and row_idx >= max_r - 2:
            is_total_row = True
        elif total_row_type == "standard" and not has_vat_summary and row_idx == max_r:
            is_total_row = True

        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.border = border_o
            
            if is_total_row:
                cell.font = font_tongket
                if total_row_type == "grand":
                    cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            else:
                cell.font = font_noidung

            # Phân loại canh lề dữ liệu theo quy chuẩn kế toán phát hành sách
            if col_idx in [1, 3]:  # STT, Mã số
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col_idx == 3: cell.number_format = '@'
            elif col_idx == 2:  # Tên sách
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_idx == 4:  # ĐVT
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:  # Các cột số tiền, chiết khấu, số lượng
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0"

    # 3. Tự động giãn rộng cột thông minh theo độ dài chữ
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

# --- KHỞI CHẠY QUY TRÌNH XỬ LÝ ---
uploaded_file = st.file_uploader("Kéo thả file Excel thô của hệ thống vào đây:", type=["xlsx"])

if uploaded_file is not None:
    st.success("Đã nhận file thô thành công!")
    
    if st.button("🚀 BẮT ĐẦU XỬ LÝ VÀ PHÂN TÍCH MASTER PROCESS"):
        try:
            with st.spinner("Hệ thống đang định vị tọa độ cột động và xử lý dữ liệu kế toán..."):
                excel_file = pd.ExcelFile(uploaded_file)
                all_sheets_data = []
                
                # --- BƯỚC 1: THUẬT TOÁN ĐỊNH VỊ TỌA ĐỘ CỘT TỰ ĐỘNG ---
                for sheet_name in excel_file.sheet_names:
                    df_raw_sheet = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                    if df_raw_sheet.empty: continue
                    
                    # Bản đồ ánh xạ để tìm đúng vị trí cột dựa trên từ khóa cốt lõi
                    col_mapping = {}
                    header_row_idx = None
                    
                    for idx, row in df_raw_sheet.iterrows():
                        row_str = row.astype(str).str.strip().tolist()
                        
                        # Quét tìm dòng chứa từ khóa mấu chốt để xác định tiêu đề
                        if any("Mã số" in s or "Tên sách" in s or "Giá bìa" in s for s in row_str):
                            header_row_idx = idx
                            for c_idx, cell_val in enumerate(row_str):
                                if "Tên sách" in cell_val or "Tên hàng" in cell_val: col_mapping['Tên sách'] = c_idx
                                elif "Mã số" in cell_val or "Mã hàng" in cell_val: col_mapping['Mã số'] = c_idx
                                elif "ĐVT" in cell_val or "Đơn vị" in cell_val: col_mapping['ĐVT'] = c_idx
                                elif "Giá bìa" in cell_val: col_mapping['Giá bìa'] = c_idx
                                elif "CK" in cell_val or "Chiết khấu" in cell_val: col_mapping['CK'] = c_idx
                                elif "Số lượng" in cell_val or "SL" in cell_val: col_mapping['Số lượng'] = c_idx
                                elif "Đơn giá" in cell_val: col_mapping['Đơn giá'] = c_idx
                                elif "Thành tiền" in cell_val: col_mapping['Thành tiền'] = c_idx
                            break
                    
                    # Yêu cầu tối thiểu phải tìm thấy 3 cột cốt lõi để nhận diện bảng dữ liệu sách hợp lệ
                    yeu_cau_cot = ['Tên sách', 'Mã số', 'Số lượng']
                    if not all(k in col_mapping for k in yeu_cau_cot) or header_row_idx is None:
                        continue  # Bỏ qua nếu sheet này là sheet trống hoặc thông tin hành chính không liên quan
                    
                    # Cắt lấy phần dữ liệu từ sau dòng tiêu đề trở đi
                    df_data = df_raw_sheet.iloc[header_row_idx + 1:].copy()
                    
                    # Tạo cấu trúc DataFrame mới dựa trên đúng tọa độ cột đã tìm được
                    df_clean = pd.DataFrame()
                    df_clean['Tên sách'] = df_data[col_mapping['Tên sách']] if 'Tên sách' in col_mapping else ""
                    df_clean['Mã số'] = df_data[col_mapping['Mã số']] if 'Mã số' in col_mapping else ""
                    df_clean['ĐVT'] = df_data[col_mapping['ĐVT']] if 'ĐVT' in col_mapping else "Cuốn"
                    df_clean['Giá bìa'] = df_data[col_mapping['Giá bìa']] if 'Giá bìa' in col_mapping else 0
                    df_clean['CK'] = df_data[col_mapping['CK']] if 'CK' in col_mapping else 0
                    df_clean['Số lượng'] = df_data[col_mapping['Số lượng']] if 'Số lượng' in col_mapping else 0
                    df_clean['Đơn giá'] = df_data[col_mapping['Đơn giá']] if 'Đơn giá' in col_mapping else 0
                    df_clean['Thành tiền'] = df_data[col_mapping['Thành tiền']] if 'Thành tiền' in col_mapping else 0
                    
                    # Làm sạch text danh mục
                    for col in ['Tên sách', 'Mã số', 'ĐVT']:
                        df_clean[col] = df_clean[col].astype(str).str.strip()
                        
                    # Bộ lọc triệt tiêu dòng rác và tổng kết phụ
                    df_clean = df_clean[df_clean['Mã số'].notna() & (df_clean['Mã số'] != '') & (df_clean['Mã số'] != 'nan') & (df_clean['Mã số'] != 'Mã số')]
                    tu_khoa_xoa = ['Tổng cộng:', 'Thuế VAT:', 'Thành tiền:', 'Tổng cộng', 'Thuế VAT', 'Phụ trách cung tiêu', 'Người giao hàng', 'Thủ Kho', 'nan', 'STT', 'Tên sách']
                    df_clean = df_clean[~df_clean['Tên sách'].isin(tu_khoa_xoa)]
                    df_clean = df_clean[~df_clean['Mã số'].isin(tu_khoa_xoa)]
                    
                    all_sheets_data.append(df_clean)
                
                if not all_sheets_data:
                    st.error("Hệ thống định vị thông minh không tìm thấy bảng dữ liệu phát hành hợp lệ. Vui lòng kiểm tra lại từ khóa tiêu đề cột trong file Excel.")
                    st.stop()
                    
                df_master = pd.concat(all_sheets_data, ignore_index=True)
                
                # Ép kiểu số cho toàn bộ mảng tính toán tài chính
                for col in ['Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']:
                    df_master[col] = pd.to_numeric(df_master[col], errors='coerce').fillna(0)
                
                # Loại bỏ dòng không phát sinh số lượng
                df_master = df_master[df_master['Số lượng'] != 0]
                df_master.insert(0, 'STT', range(1, len(df_master) + 1))
                
                # --- BƯỚC 2: PHÂN TÍCH VÀ GOM NHÓM DANH MỤC (LOGIC FILE .BAS) ---
                df_duong = df_master[df_master['Số lượng'] >= 0].copy()
                df_am = df_master[df_master['Số lượng'] < 0].copy()
                
                # 2.1 Tổng hợp mã hàng bán (Gom theo Mã số, đơn giá bìa gốc)
                df_tonghop_ban = df_duong.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa'], as_index=False).agg({'Số lượng': 'sum'})
                df_tonghop_ban['CK'] = 0
                df_tonghop_ban['Đơn giá'] = df_tonghop_ban['Giá bìa']
                df_tonghop_ban['Thành tiền'] = df_tonghop_ban['Số lượng'] * df_tonghop_ban['Đơn giá']
                df_tonghop_ban = df_tonghop_ban[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                df_tonghop_ban.insert(0, 'STT', range(1, len(df_tonghop_ban) + 1))
                
                # 2.2 Tổng hợp hàng trả (Chuyển dấu âm sang dương để làm báo cáo xuất nhập)
                df_tonghop_tra = df_am.copy()
                df_tonghop_tra['Số lượng'] = df_tonghop_tra['Số lượng'].abs()
                df_tonghop_tra = df_tonghop_tra.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa'], as_index=False).agg({'Số lượng': 'sum'})
                df_tonghop_tra['CK'] = 0
                df_tonghop_tra['Đơn giá'] = df_tonghop_tra['Giá bìa']
                df_tonghop_tra['Thành tiền'] = df_tonghop_tra['Số lượng'] * df_tonghop_tra['Đơn giá']
                df_tonghop_tra = df_tonghop_tra[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                df_tonghop_tra.insert(0, 'STT', range(1, len(df_tonghop_tra) + 1))
                
                # 2.3 Chi tiết hàng trả âm (Giữ nguyên tỷ lệ chiết khấu và dấu âm thực tế)
                df_chitiet_am = df_am.copy()
                df_chitiet_am['Thành tiền'] = df_chitiet_am['Số lượng'] * df_chitiet_am['Đơn giá']
                df_chitiet_am.insert(0, 'STT', range(1, len(df_chitiet_am) + 1))
                df_chitiet_am = df_chitiet_am[['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                
                # 2.4 Toàn bộ danh mục xuất bán dương
                df_tonghop_banra = df_duong.copy()
                df_tonghop_banra['Thành tiền'] = df_tonghop_banra['Số lượng'] * df_tonghop_banra['Đơn giá']
                df_tonghop_banra.insert(0, 'STT', range(1, len(df_tonghop_banra) + 1))
                df_tonghop_banra = df_tonghop_banra[['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]

                # --- BƯỚC 3: ĐÓNG GÓI VÀ XUẤT BÁO CÁO EXCEL CHUYÊN NGHIỆP ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    grand_summary_data = []
                    
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

                    # Xuất bản dữ liệu ra các sheet chức năng theo mã màu
                    if not df_tonghop_ban.empty:
                        ln, sl, tt, _, _ = ghi_sheet_kem_tong(df_tonghop_ban, "Tong_Hop_Ma_Hang_Ban", "003366")
                        grand_summary_data.append(["Tong Hop Ma Hang Ban", "Duong", ln, sl, tt, 0, tt, "Gia Bia (Khong VAT)"])
                        
                    if not df_tonghop_tra.empty:
                        ln, sl, tt, _, _ = ghi_sheet_kem_tong(df_tonghop_tra, "Tong_Hop_Hang_Tra", "FF6600")
                        grand_summary_data.append(["Tong Hop Hang Tra", "Duong", ln, sl, tt, 0, tt, "Gia Bia (Khong VAT)"])
                        
                    if not df_chitiet_am.empty:
                        ln, sl, tt, vt, st = ghi_sheet_kem_tong(df_chitiet_am, "Chi_Tiet_Am", "C00000", has_vat=True)
                        grand_summary_data.append(["Chi Tiet Hang Tra", "Am", ln, sl, tt, vt, st, "Thuc te sau CK"])
                        
                    if not df_tonghop_banra.empty:
                        ln, sl, tt, vt, st = ghi_sheet_kem_tong(df_tonghop_banra, "Tong_Hop_Ban_Ra", "00B050", has_vat=True)
                        grand_summary_data.append(["Tong Hop Ban Ra", "Duong", ln, sl, tt, vt, st, "Toan bo du lieu duong"])

                    # Phân chia nhỏ hóa đơn để đẩy lên hệ thống hóa đơn điện tử (Batch = 1000 dòng)
                    batch_size = 1000
                    sheet_index = 1
                    for i in range(0, len(df_tonghop_banra), batch_size):
                        df_batch = df_tonghop_banra.iloc[i:i+batch_size].copy()
                        df_batch['STT'] = range(1, len(df_batch) + 1)
                        hd_name = f"HD {sheet_index}"
                        ln, sl, tt, vt, st = ghi_sheet_kem_tong(df_batch, hd_name, "0070C0", has_vat=True)
                        grand_summary_data.append([f"Hoa Don {sheet_index}", "Duong", ln, sl, tt, vt, st, "Tach le 1000 dong"])
                        sheet_index += 1

                    # Xuất báo cáo tài chính tổng quan "Tong_Ket_Chung"
                    df_grand = pd.DataFrame(grand_summary_data, columns=["Ten Sheet / Hang muc", "Loai", "So dong", "So luong", "Truoc Thue", "VAT (5%)", "Sau Thue", "Ghi chu"])
                    df_grand.to_excel(writer, index=False, sheet_name="Tong_Ket_Chung")
                    ws_grand = writer.sheets["Tong_Ket_Chung"]
                    
                    # Doanh thu thuần thực tế bù trừ tuyệt đối của doanh nghiệp sau phát hành
                    thuong_sl = df_tonghop_banra['Số lượng'].sum() + df_chitiet_am['Số lượng'].sum()
                    thuong_tt = df_tonghop_banra['Thành tiền'].sum() + df_chitiet_am['Thành tiền'].sum()
                    thuong_vat = thuong_tt * 0.05
                    thuong_st = thuong_tt * 1.05
                    
                    r_grand = len(df_grand) + 2
                    ws_grand.cell(row=r_grand, column=1).value = "DOANH THU THUC TE (BAN - TRA)"
                    ws_grand.cell(row=r_grand, column=4).value = thuong_sl
                    ws_grand.cell(row=r_grand, column=5).value = thuong_tt
                    ws_grand.cell(row=r_grand, column=6).value = thuong_vat
                    ws_grand.cell(row=r_grand, column=7).value = thuong_st
                    
                    trang_tri_sheet(ws_grand, "000000", total_row_type="grand")
                    
                    # Đưa sheet tổng quan tài chính lên đầu danh sách sheet
                    wb = writer.book
                    order = [wb.sheetnames[-1]] + wb.sheetnames[:-1]
                    wb._sheets = [wb._sheets[wb.sheetnames.index(name)] for name in order]

                processed_data = output.getvalue()
                st.success("🎉 PHÂN TÍCH MASTER PROCESS THÀNH CÔNG HOÀN HẢO!")
                
                # --- PHẢN HỒI TAB TRỰC QUAN LÊN INTERFACE WEB ---
                tab1, tab2, tab3, tab4 = st.tabs(["📋 Tổng Kết Chung", "📈 Tổng Hợp Mã Bán", "📉 Tổng Hợp Mã Trả", "📦 Chi Tiết Hàng Âm"])
                with tab1:
                    st.dataframe(df_grand, use_container_width=True)
                    st.metric(label="Doanh Thu Thuần Thực Tế Của Hệ Thống (Sau Thuế)", value=f"{thuong_st:,.0f} đ")
                with tab2:
                    st.dataframe(df_tonghop_ban, use_container_width=True)
                with tab3:
                    st.dataframe(df_tonghop_tra, use_container_width=True)
                with tab4:
                    st.dataframe(df_chitiet_am, use_container_width=True)

                # Download Button
                st.download_button(
                    label="📥 TẢI FILE EXCEL MASTER REPORT (FULL SHEETS)",
                    data=processed_data,
                    file_name=f"Master_Report_Sach_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Hệ thống gặp sự cố nghiêm trọng: {str(e)}")
