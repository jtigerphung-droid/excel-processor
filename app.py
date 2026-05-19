import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- GIAO DIỆN HỆ THỐNG ---
st.set_page_config(page_title="Hệ thống Master Process Phát Hành", layout="wide")
st.title("📊 HỆ THỐNG XỬ LÝ DỮ LIỆU PHÁT HÀNH SÁCH TỰ ĐỘNG")
st.write("Phiên bản Tối ưu hóa Logic: Tự động bóc tách dữ liệu lõi, cô lập dòng rác hành chính và xử lý bù trừ âm dương.")

# --- HÀM TRANG TRÍ EXCEL THEO QUY CHUẨN KẾ TOÁN ---
def trang_tri_sheet(worksheet, tieude_color, has_vat_summary=False, total_row_type="standard"):
    font_tieude = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fill_tieude = PatternFill(start_color=tieude_color, end_color=tieude_color, fill_type="solid")
    font_noidung = Font(name="Arial", size=10)
    font_tongket = Font(name="Arial", size=10, bold=True)
    
    vien_mong = Side(border_style="thin", color="D9D9D9")
    border_o = Border(left=vien_mong, right=vien_mong, top=vien_mong, bottom=vien_mong)
    
    # 1. Tiêu đề cột
    worksheet.row_dimensions[1].height = 25
    for col_idx in range(1, worksheet.max_column + 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.font = font_tieude
        cell.fill = fill_tieude
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_o

    # 2. Định dạng nội dung dữ liệu sách
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

            # Định dạng hiển thị và canh lề cột dữ liệu
            if col_idx in [1, 3]:  # STT, Mã số sách
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col_idx == 3: cell.number_format = '@'
            elif col_idx == 2:  # Tên sách
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_idx == 4:  # ĐVT
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:  # Các cột tài chính: Giá bìa, Chiết khấu, Số lượng, Đơn giá, Thành tiền
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0"

    # 3. Tự động giãn cột linh hoạt
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

# --- XỬ LÝ TỆP TIN TẢI LÊN ---
uploaded_file = st.file_uploader("Kéo thả file Excel thô của hệ thống vào đây:", type=["xlsx"])

if uploaded_file is not None:
    st.success("Đã nhận file Excel thành công!")
    
    if st.button("🚀 BẮT ĐẦU XỬ LÝ VÀ PHÂN TÍCH MASTER PROCESS"):
        try:
            with st.spinner("Đang triển khai thuật toán quét lọc diện rộng và đồng bộ dữ liệu lõi..."):
                excel_file = pd.ExcelFile(uploaded_file)
                all_cleaned_rows = []
                
                # --- BƯỚC 1: ĐỌC THÔ DIỆN RỘNG VÀ TRÍCH XUẤT DỮ LIỆU LÕI ---
                for sheet_name in excel_file.sheet_names:
                    df_raw_sheet = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                    if df_raw_sheet.empty: continue
                    
                    # Quét từng dòng của sheet để tìm các dòng chứa dữ liệu sách thực sự
                    for idx, row in df_raw_sheet.iterrows():
                        row_vals = row.fillna("").astype(str).str.strip().tolist()
                        
                        # Bỏ qua các dòng tiêu đề, dòng trống hoặc dòng văn bản hành chính tổng cộng
                        if not row_vals or len(row_vals) < 7: continue
                        if any(k in "".join(row_vals) for k in ["Tổng cộng", "Thuế VAT", "Thành tiền", "Người lập", "Thủ kho", "Giám đốc"]): continue
                        if "Mã số" in row_vals or "Tên sách" in row_vals or "STT" in row_vals: continue
                        
                        # THUẬT TOÁN ĐỊNH VỊ ĐỘNG: Tìm vị trí cột dựa trên định dạng dữ liệu trong dòng
                        # Tìm ô chứa mã số (thường là chuỗi ký tự số dài hoặc mã sách đặc thù)
                        # Để an toàn và đồng bộ, ta dò tìm cấu trúc phân bổ cột chuẩn hóa 9 cột
                        idx_ma = -1
                        idx_ten = -1
                        idx_sl = -1
                        
                        for c_i, val in enumerate(row_vals):
                            # Ô chứa mã số thường là chuỗi ký tự không chứa khoảng trắng và có độ dài phù hợp
                            if val and not " " in val and len(val) >= 4 and idx_ma == -1:
                                # Kiểm tra xem đây có phải mã hàng/mã số không
                                if re.match(r'^[A-Za-z0-9\-_.]+$', val): idx_ma = c_i
                            # Ô chứa tên sách thường là chuỗi dài có khoảng trắng
                            elif val and " " in val and len(val) > 10 and idx_ten == -1:
                                idx_ten = c_i
                        
                        # Nếu dòng tính toán này không nhận diện được cấu trúc cơ bản, xử lý theo chỉ mục mặc định (Fallback)
                        if idx_ma == -1 or idx_ten == -1:
                            # Nếu dòng có cột 1 là số thứ tự, lấy cấu trúc mặc định chuẩn hệ thống phát hành (A->I)
                            if row_vals[0].isdigit() or (row_vals[2] != "" and len(row_vals[2]) >= 4):
                                idx_ma = 2
                                idx_ten = 1
                            else:
                                continue # Dòng rác, bỏ qua
                                
                        # Trích xuất dữ liệu theo đúng sơ đồ phân bổ cột phát hành
                        try:
                            # Khôi phục chỉ mục tương đối quanh cột Tên sách và Mã số
                            base_idx = min(idx_ten, idx_ma)
                            if base_idx > 0 and row_vals[base_idx-1].isdigit():
                                stt_val = row_vals[base_idx-1]
                            else:
                                stt_val = ""
                                
                            ten_sach = row_vals[idx_ten]
                            ma_so = row_vals[idx_ma]
                            dvt = row_vals[idx_ma + 1] if (idx_ma + 1) < len(row_vals) else "Cuốn"
                            
                            # Thu thập các trường số liệu tài chính ở phía sau
                            num_fields = []
                            for v in row_vals[idx_ma + 2:]:
                                if v == "" or v == "nan": num_fields.append(0)
                                else:
                                    # Làm sạch các ký tự phân tách hàng ngàn như dấu phẩy hoặc dấu chấm
                                    v_clean = v.replace(",", "").replace(".", "").strip()
                                    if v_clean.replace("-", "").isdigit(): num_fields.append(float(v_clean))
                                    else: num_fields.append(0)
                                    
                            # Đảm bảo mảng số liệu điền đầy đủ cho các cột tài chính
                            while len(num_fields) < 5: num_fields.append(0)
                            
                            gia_bia = num_fields[0]
                            ck = num_fields[1]
                            sl = num_fields[2]
                            don_gia = num_fields[3]
                            thanh_tien = num_fields[4]
                            
                            # Nếu số lượng bằng 0 hoặc dữ liệu mã rác, loại bỏ lập tức
                            if sl == 0 or ma_so == "" or ma_so == "nan": continue
                            
                            all_cleaned_rows.append([
                                stt_val, ten_sach, ma_so, dvt, gia_bia, ck, sl, don_gia, thanh_tien
                            ])
                        except Exception:
                            continue # Dòng lỗi cấu trúc cục bộ, bỏ qua để bảo vệ luồng chính

                if not all_cleaned_rows:
                    st.error("❌ LỖI HỆ THỐNG: Bộ lọc lõi diện rộng không trích xuất được dòng dữ liệu sách hợp lệ nào. Vui lòng kiểm tra lại file Excel đầu vào.")
                    st.stop()
                    
                # Tạo DataFrame tổng Master chính xác tuyệt đối từ mảng dữ liệu sạch
                df_master = pd.DataFrame(all_cleaned_rows, columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                df_master['STT'] = range(1, len(df_master) + 1)
                
                # --- BƯỚC 2: TRIỂN KHAI THUẬT TOÁN DICTIONARY (LOGIC FILE .BAS CHUẨN) ---
                df_duong = df_master[df_master['Số lượng'] >= 0].copy()
                df_am = df_master[df_master['Số lượng'] < 0].copy()
                
                # 2.1 Tổng hợp mã hàng bán (Gom theo Mã số, Đơn giá = Giá bìa gốc)
                if not df_duong.empty:
                    df_tonghop_ban = df_duong.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa'], as_index=False).agg({'Số lượng': 'sum'})
                    df_tonghop_ban['CK'] = 0
                    df_tonghop_ban['Đơn giá'] = df_tonghop_ban['Giá bìa']
                    df_tonghop_ban['Thành tiền'] = df_tonghop_ban['Số lượng'] * df_tonghop_ban['Đơn giá']
                    df_tonghop_ban = df_tonghop_ban[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                    df_tonghop_ban.insert(0, 'STT', range(1, len(df_tonghop_ban) + 1))
                else:
                    df_tonghop_ban = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                
                # 2.2 Tổng hợp hàng trả (Đảo dấu số lượng từ Âm sang Dương - Tính theo Giá bìa gốc)
                if not df_am.empty:
                    df_tonghop_tra = df_am.copy()
                    df_tonghop_tra['Số lượng'] = df_tonghop_tra['Số lượng'].abs()
                    df_tonghop_tra = df_tonghop_tra.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa'], as_index=False).agg({'Số lượng': 'sum'})
                    df_tonghop_tra['CK'] = 0
                    df_tonghop_tra['Đơn giá'] = df_tonghop_tra['Giá bìa']
                    df_tonghop_tra['Thành tiền'] = df_tonghop_tra['Số lượng'] * df_tonghop_tra['Đơn giá']
                    df_tonghop_tra = df_tonghop_tra[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                    df_tonghop_tra.insert(0, 'STT', range(1, len(df_tonghop_tra) + 1))
                    
                    # 2.3 Chi tiết hàng trả âm (Giữ nguyên dấu âm và chiết khấu gốc của hệ thống để đối chiếu)
                    df_chitiet_am = df_am.copy()
                    df_chitiet_am['Thành tiền'] = df_chitiet_am['Số lượng'] * df_chitiet_am['Đơn giá']
                    df_chitiet_am['STT'] = range(1, len(df_chitiet_am) + 1)
                    df_chitiet_am = df_chitiet_am[['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                else:
                    df_tonghop_tra = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                    df_chitiet_am = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                
                # 2.4 Toàn bộ dữ liệu xuất bán dương thực tế
                df_tonghop_banra = df_duong.copy()
                if not df_tonghop_banra.empty:
                    df_tonghop_banra['Thành tiền'] = df_tonghop_banra['Số lượng'] * df_tonghop_banra['Đơn giá']
                    df_tonghop_banra['STT'] = range(1, len(df_tonghop_banra) + 1)
                    df_tonghop_banra = df_tonghop_banra[['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                else:
                    df_tonghop_banra = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])

                # --- BƯỚC 3: ĐÓNG GÓI VÀ XUẤT FILE EXCEL ĐA SHEET ---
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

                    # Ghi các báo cáo chuyên biệt vào các sheet theo đúng gam màu quản trị
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

                    # Tiến hành thực hiện phân tách Hóa đơn (Mỗi hóa đơn tối đa 1000 dòng)
                    batch_size = 1000
                    sheet_index = 1
                    if not df_tonghop_banra.empty:
                        for i in range(0, len(df_tonghop_banra), batch_size):
                            df_batch = df_tonghop_banra.iloc[i:i+batch_size].copy()
                            df_batch['STT'] = range(1, len(df_batch) + 1)
                            hd_name = f"HD {sheet_index}"
                            ln, sl, tt, vt, st = ghi_sheet_kem_tong(df_batch, hd_name, "0070C0", has_vat=True)
                            grand_summary_data.append([f"Hoa Don {sheet_index}", "Duong", ln, sl, tt, vt, st, "Tach le 1000 dong"])
                            sheet_index += 1

                    # Tạo trang Dashboard Báo cáo Doanh thu "Tong_Ket_Chung" đưa lên đầu file
                    df_grand = pd.DataFrame(grand_summary_data, columns=["Ten Sheet / Hang muc", "Loai", "So dong", "So luong", "Truoc Thue", "VAT (5%)", "Sau Thue", "Ghi chu"])
                    df_grand.to_excel(writer, index=False, sheet_name="Tong_Ket_Chung")
                    ws_grand = writer.sheets["Tong_Ket_Chung"]
                    
                    # Công thức tính doanh thu thuần thực tế sau bù trừ xuất - trả toàn hệ thống
                    thuong_sl = df_tonghop_banra['Số lượng'].sum() + df_chitiet_am['Số lượng'].sum() if not df_chitiet_am.empty else df_tonghop_banra['Số lượng'].sum()
                    thuong_tt = df_tonghop_banra['Thành tiền'].sum() + df_chitiet_am['Thành tiền'].sum() if not df_chitiet_am.empty else df_tonghop_banra['Thành tiền'].sum()
                    thuong_vat = thuong_tt * 0.05
                    thuong_st = thuong_tt * 1.05
                    
                    r_grand = len(df_grand) + 2
                    ws_grand.cell(row=r_grand, column=1).value = "DOANH THU THUC TE (BAN - TRA)"
                    ws_grand.cell(row=r_grand, column=4).value = thuong_sl
                    ws_grand.cell(row=r_grand, column=5).value = thuong_tt
                    ws_grand.cell(row=r_grand, column=6).value = thuong_vat
                    ws_grand.cell(row=r_grand, column=7).value = thuong_st
                    
                    trang_tri_sheet(ws_grand, "000000", total_row_type="grand")
                    
                    # Đưa trang tổng kết lên vị trí tab đầu tiên trong Workbook
                    wb = writer.book
                    order = [wb.sheetnames[-1]] + wb.sheetnames[:-1]
                    wb._sheets = [wb._sheets[wb.sheetnames.index(name)] for name in order]

                processed_data = output.getvalue()
                st.success("🎉 MASTER PROCESS HOÀN THÀNH TOÀN DIỆN!")
                
                # --- PHẢN HỒI GIAO DIỆN WEB TRỰC QUAN ---
                tab1, tab2, tab3, tab4 = st.tabs(["📋 Tổng Kết Chung (Báo cáo)", "📈 Tổng Hợp Mã Bán", "📉 Tổng Hợp Mã Trả", "📦 Chi Tiết Hàng Âm"])
                with tab1:
                    st.dataframe(df_grand, use_container_width=True)
                    st.metric(label="Doanh Thu Thuần Thực Tế Đối Soát (Sau Thuế 5%)", value=f"{thuong_st:,.0f} đ")
                with tab2:
                    st.dataframe(df_tonghop_ban, use_container_width=True)
                with tab3:
                    st.dataframe(df_tonghop_tra, use_container_width=True)
                with tab4:
                    st.dataframe(df_chitiet_am, use_container_width=True)

                    st.download_button(
                    label="📥 TẢI FILE EXCEL MASTER REPORT (FULL 100%)",
                    data=processed_data,
                    file_name=f"Master_Report_Chuan_Vi__{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ Sự cố logic nghiêm trọng trong nhân xử lý: {str(e)}")
