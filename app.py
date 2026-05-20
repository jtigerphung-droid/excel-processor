import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- CẤU HÌNH GIAO DIỆN HỆ THỐNG ---
st.set_page_config(page_title="Hệ thống Phát Hành Sách 2 Giai Đoạn", layout="wide")
st.title("📊 HỆ THỐNG XỬ LÝ DỮ LIỆU PHÁT HÀNH SÁCH")
st.write("Phiên bản kiến trúc PAB21 - Bảo toàn nguyên bản dữ liệu gốc & Tọa độ ô tĩnh.")

# --- HÀM TRANG TRÍ EXCEL THEO QUY CHUẨN KẾ TOÁN ---
def trang_tri_sheet(worksheet, tieude_color, has_vat_summary=False, total_row_type="standard"):
    font_tieude = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fill_tieude = PatternFill(start_color=tieude_color, end_color=tieude_color, fill_type="solid")
    font_noidung = Font(name="Arial", size=10)
    font_tongket = Font(name="Arial", size=10, bold=True)
    
    vien_mong = Side(border_style="thin", color="D9D9D9")
    border_o = Border(left=vien_mong, right=vien_mong, top=vien_mong, bottom=vien_mong)
    
    worksheet.row_dimensions[1].height = 25
    for col_idx in range(1, worksheet.max_column + 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.font = font_tieude
        cell.fill = fill_tieude
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_o

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

            if col_idx in [1, 3]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col_idx == 3: cell.number_format = '@'
            elif col_idx == 2:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_idx == 4:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 6:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                # Nếu là sheet tổng kết chung, cột số dòng/số lượng định dạng số nguyên, trước thuế định dạng tiền tệ
                if total_row_type == "grand":
                    cell.number_format = "#,##0"
                else:
                    cell.number_format = "0.00"  # Giữ nguyên định dạng số thô chiết khấu
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0"

    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)


# --- ĐIỀU HƯỚNG TÍNH NĂNG QUA TAB ---
tab_giai_doan_1, tab_giai_doan_2 = st.tabs(["🔄 GIAI ĐOẠN 1: Gộp & Làm Sạch (Bảo Lưu Gốc)", "🧮 GIAI ĐOẠN 2: Tính Toán Phân Tách PAB21"])

# ==========================================================================================
# GIAI ĐOẠN 1: BẢO LƯU NGUYÊN BẢN 100% CODE GỐC CHẠY HOÀN CHỈNH CỦA USER
# ==========================================================================================
with tab_giai_doan_1:
    st.header("Bước 1: Gộp Nhiều Sheet & Làm Sạch Diện Rộng")
    st.info("Mã nguồn phần này được bảo lưu nguyên vẹn theo cấu trúc chạy ổn định trước đó.")
    
    uploaded_file = st.file_uploader("Kéo thả file Excel thô của hệ thống vào đây:", type=["xlsx"], key="file_tho_goc")

    if uploaded_file is not None:
        if st.button("🚀 THỰC HIỆN GỘP VÀ LÀM SẠCH PHIÊN BẢN CHUẨN"):
            try:
                excel_file = pd.ExcelFile(uploaded_file)
                all_cleaned_rows = []
                
                for sheet_name in excel_file.sheet_names:
                    df_raw_sheet = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                    if df_raw_sheet.empty: continue
                    
                    for idx, row in df_raw_sheet.iterrows():
                        row_vals = row.fillna("").astype(str).str.strip().tolist()
                        
                        if not row_vals or len(row_vals) < 7: continue
                        if any(k in "".join(row_vals) for k in ["Tổng cộng", "Thuế VAT", "Thành tiền", "Người lập", "Thủ kho", "Giám đốc"]): continue
                        if "Mã số" in row_vals or "Tên sách" in row_vals or "STT" in row_vals: continue
                        
                        idx_ma = -1
                        idx_ten = -1
                        
                        for c_i, val in enumerate(row_vals):
                            if val and not " " in val and len(val) >= 4 and idx_ma == -1:
                                if re.match(r'^[A-Za-z0-9\-_.]+$', val): idx_ma = c_i
                            elif val and " " in val and len(val) > 10 and idx_ten == -1:
                                idx_ten = c_i
                        
                        if idx_ma == -1 or idx_ten == -1:
                            if row_vals[0].isdigit() or (row_vals[2] != "" and len(row_vals[2]) >= 4):
                                idx_ma = 2
                                idx_ten = 1
                            else:
                                continue
                                
                        try:
                            base_idx = min(idx_ten, idx_ma)
                            if base_idx > 0 and row_vals[base_idx-1].isdigit():
                                stt_val = row_vals[base_idx-1]
                            else:
                                stt_val = ""
                                
                            ten_sach = row_vals[idx_ten]
                            ma_so = row_vals[idx_ma]
                            dvt = row_vals[idx_ma + 1] if (idx_ma + 1) < len(row_vals) else "Cuốn"
                            
                            num_fields = []
                            for v in row_vals[idx_ma + 2:]:
                                if v == "" or v == "nan": num_fields.append(0)
                                else:
                                    v_clean = v.replace(",", "").replace(".", "").strip()
                                    if v_clean.replace("-", "").isdigit(): num_fields.append(float(v_clean))
                                    else: num_fields.append(0)
                                    
                            while len(num_fields) < 5: num_fields.append(0)
                            
                            gia_bia = num_fields[0]
                            ck = num_fields[1]
                            sl = num_fields[2]
                            don_gia = num_fields[3]
                            thanh_tien = num_fields[4]
                            
                            if sl == 0 or ma_so == "" or ma_so == "nan": continue
                            
                            all_cleaned_rows.append([
                                stt_val, ten_sach, ma_so, dvt, gia_bia, ck, sl, don_gia, thanh_tien
                            ])
                        except Exception:
                            continue

                if not all_cleaned_rows:
                    st.error("❌ LỖI: Không trích xuất được dòng dữ liệu sách hợp lệ nào.")
                    st.stop()
                    
                df_master = pd.DataFrame(all_cleaned_rows, columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                df_master['STT'] = range(1, len(df_master) + 1)
                
                out_sach = io.BytesIO()
                with pd.ExcelWriter(out_sach, engine='openpyxl') as writer:
                    df_master.to_excel(writer, index=False, sheet_name="Du_Lieu_Sach_100")
                    trang_tri_sheet(writer.sheets["Du_Lieu_Sach_100"], "003366")
                    
                st.success("🎉 GIAI ĐOẠN 1 HOÀN THÀNH: XUẤT FILE 9 CỘT CHUẨN!")
                st.dataframe(df_master, use_container_width=True)
                
                st.download_button(
                    label="📥 TẢI FILE DỮ LIỆU SẠCH (Dùng cho Giai đoạn 2)",
                    data=out_sach.getvalue(),
                    file_name=f"DuLieu_Gop_Va_LamSach_Chuan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Lỗi hệ thống G1: {str(e)}")

# ==========================================================================================
# GIAI ĐOẠN 2: KIẾN TRÚC MÃ NGUỒN THEO PHƯƠNG PHÁP PAB21 (TỌA ĐỘ Ô TUYỆT ĐỐI)
# ==========================================================================================
with tab_giai_doan_2:
    st.header("Bước 2: Triển Khai Thuật Toán Đối Soát Mảng Phẳng PAB21")
    st.warning("⚠️ ĐIỀU KIỆN: Nạp chính xác file Excel kết quả 'Du_Lieu_Sach_100' từ Giai đoạn 1.")
    
    file_sach = st.file_uploader("Tải lên file Excel ĐÃ LÀM SẠCH CHUẨN:", type=["xlsx"], key="file_sach_pab21")
    
    if file_sach is not None:
        if st.button("🧮 KHỞI CHẠY KIẾN TRÚC TÍNH TOÁN PAB21"):
            try:
                # DÙNG OPENPYXL ĐỂ ĐỌC Ô TĨNH TUYỆT ĐỐI - TRIỆT TIÊU HOÀN TOÀN LỖI PANDAS TỰ ĐOÁN
                wb_in = openpyxl.load_workbook(file_sach, data_only=True)
                if "Du_Lieu_Sach_100" not in wb_in.sheetnames:
                    st.error("❌ LỖI: File nạp vào không chứa sheet 'Du_Lieu_Sach_100' tiêu chuẩn.")
                    st.stop()
                    
                ws_in = wb_in["Du_Lieu_Sach_100"]
                
                mảng_bán_dương = []
                mảng_trả_âm = []
                
                # Duyệt tuyến tính qua mảng phẳng từ dòng số 2 đến hết bảng tính
                for row_idx in range(2, ws_in.max_row + 1):
                    # Đọc dữ liệu thô theo đúng định vị cột cố định từ 1 đến 9 (A đến I)
                    val_stt   = ws_in.cell(row=row_idx, column=1).value
                    val_ten   = str(ws_in.cell(row=row_idx, column=2).value or "").strip()
                    val_ma    = str(ws_in.cell(row=row_idx, column=3).value or "").strip()
                    val_dvt   = str(ws_in.cell(row=row_idx, column=4).value or "Cuốn").strip()
                    
                    try: val_bia = float(ws_in.cell(row=row_idx, column=5).value or 0)
                    except: val_bia = 0.0
                    
                    try: val_ck = float(ws_in.cell(row=row_idx, column=6).value or 0)
                    except: val_ck = 0.0
                        
                    try: val_sl = float(ws_in.cell(row=row_idx, column=7).value or 0)
                    except: val_sl = 0.0
                        
                    try: val_dg = float(ws_in.cell(row=row_idx, column=8).value or 0)
                    except: val_dg = 0.0
                        
                    try: val_tt = float(ws_in.cell(row=row_idx, column=9).value or 0)
                    except: val_tt = 0.0
                    
                    # Bỏ qua dòng trống hoặc dòng rác không có mã hàng
                    if not val_ma or val_ma == "None" or val_sl == 0:
                        continue
                        
                    dòng_dữ_liệu = [val_stt, val_ten, val_ma, val_dvt, val_bia, val_ck, val_sl, val_dg, val_tt]
                    
                    # BỘ LỌC BƯỚC ĐỆM PAB21: PHÂN LOẠI ĐỘC LẬP THEO DẤU SỐ LƯỢNG
                    if val_sl > 0:
                        mảng_bán_dương.append(dòng_dữ_liệu)
                    else:
                        # BẢO TOÀN DẤU ÂM: Giữ nguyên toàn bộ dấu âm của dữ liệu gốc, không chuyển dương
                        mảng_trả_âm.append(dòng_dữ_liệu)

                # --- 1. PHÉP TÍNH SHEET TỔNG HỢP HÀNG BÁN THEO GIÁ BÌA (Mảng Dương, Lọc trùng theo Mã) ---
                dic_gia_bia = {}
                for r in mảng_bán_dương:
                    ma_hang = r[2]
                    ten_hang = r[1]
                    dvt_hang = r[3]
                    gia_bia = r[4]
                    sl_hang = r[6]
                    
                    if ma_hang not in dic_gia_bia:
                        dic_gia_bia[ma_hang] = {"ten": ten_hang, "dvt": dvt_hang, "bia": gia_bia, "sl": 0.0}
                    dic_gia_bia[ma_hang]["sl"] += sl_hang

                bảng_gia_bia = []
                for idx, (ma, item) in enumerate(dic_gia_bia.items(), 1):
                    tt_bia = item["sl"] * item["bia"]
                    bảng_gia_bia.append([idx, item["ten"], ma, item["dvt"], item["bia"], 0.0, item["sl"], item["bia"], tt_bia])

                # --- 2. PHÉP TÍNH SHEET TỔNG HỢP HÀNG BÁN THEO CHIẾT KHẤU (Mảng Dương, Lọc trùng Mã + CK, Bảo toàn Đơn giá gốc) ---
                dic_chiet_khau = {}
                for r in mảng_bán_dương:
                    ma_hang = r[2]
                    ten_hang = r[1]
                    dvt_hang = r[3]
                    gia_bia = r[4]
                    ck_hang = r[5]  # Giữ nguyên định dạng số gốc (Ví dụ 0.3 hoặc 30)
                    sl_hang = r[6]
                    dg_goc = r[7]   # Bốc nguyên văn đơn giá gốc từ file thô sang
                    
                    key_composite = f"{ma_hang}_{ck_hang}"
                    if key_composite not in dic_chiet_khau:
                        dic_chiet_khau[key_composite] = {
                            "ten": ten_hang, "ma": ma_hang, "dvt": dvt_hang, 
                            "bia": gia_bia, "ck": ck_hang, "sl": 0.0, "dg": dg_goc
                        }
                    dic_chiet_khau[key_composite]["sl"] += sl_hang

                bảng_chiet_khau = []
                for idx, (k, item) in enumerate(dic_chiet_khau.items(), 1):
                    tt_ck = item["sl"] * item["dg"]
                    bảng_chiet_khau.append([idx, item["ten"], item["ma"], item["dvt"], item["bia"], item["ck"], item["sl"], item["dg"], tt_ck])

                # --- 3. PHÉP TÍNH SHEET TỔNG HỢP HÀNG TRẢ (Mảng Âm, Lọc trùng Mã + CK, Bảo toàn dấu âm) ---
                dic_hang_tra = {}
                for r in mảng_trả_âm:
                    ma_hang = r[2]
                    ten_hang = r[1]
                    dvt_hang = r[3]
                    gia_bia = r[4]
                    ck_hang = r[5]
                    sl_hang = r[6] # Mang dấu âm (-)
                    dg_goc = r[7]  # Đơn giá thô gốc (dương)
                    
                    key_composite = f"{ma_hang}_{ck_hang}"
                    if key_composite not in dic_hang_tra:
                        dic_hang_tra[key_composite] = {
                            "ten": ten_hang, "ma": ma_hang, "dvt": dvt_hang, 
                            "bia": gia_bia, "ck": ck_hang, "sl": 0.0, "dg": dg_goc
                        }
                    dic_hang_tra[key_composite]["sl"] += sl_hang # Cộng dồn đại số giữ nguyên dấu âm

                bảng_hang_tra = []
                for idx, (k, item) in enumerate(dic_hang_tra.items(), 1):
                    tt_tra = item["sl"] * item["dg"] # Số lượng âm nhân đơn giá dương tự động ra thành tiền âm
                    bảng_hang_tra.append([idx, item["ten"], item["ma"], item["dvt"], item["bia"], item["ck"], item["sl"], item["dg"], tt_tra])

                # --- KHỞI TẠO TIẾN TRÌNH ĐÓNG GÓI RA FILE EXCEL ĐA SHEET MÀU SẮC PAB21 ---
                out_report = io.BytesIO()
                columns_standard = ['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']
                
                with pd.ExcelWriter(out_report, engine='openpyxl') as writer:
                    grand_summary_list = []
                    
                    def ghi_sheet_he_thong(data_list, sheet_name, header_color, is_data_thô=False, has_vat=False):
                        df_sheet = pd.DataFrame(data_list, columns=columns_standard)
                        df_sheet.to_excel(writer, index=False, sheet_name=sheet_name)
                        ws = writer.sheets[sheet_name]
                        r_total = len(df_sheet) + 2
                        
                        sum_sl = df_sheet['Số lượng'].sum()
                        sum_tt = df_sheet['Thành tiền'].sum()
                        
                        # Tạo dòng Tổng kết cuối bảng tính
                        ws.cell(row=r_total, column=6).value = "Tổng cộng:"
                        ws.cell(row=r_total, column=7).value = sum_sl
                        ws.cell(row=r_total, column=9).value = sum_tt
                        
                        vat_val, sau_thue_val = 0.0, 0.0
                        if has_vat:
                            vat_val = sum_tt * 0.05
                            sau_thue_val = sum_tt + vat_val
                            ws.cell(row=r_total+1, column=8).value = "VAT 5%:"
                            ws.cell(row=r_total+1, column=9).value = vat_val
                            ws.cell(row=r_total+2, column=8).value = "Sau Thuế:"
                            ws.cell(row=r_total+2, column=9).value = sau_thue_val
                            
                        trang_tri_sheet(ws, header_color, has_vat_summary=has_vat)
                        return len(df_sheet), sum_sl, sum_tt, vat_val, sau_thue_val

                    # A. GHI HAI SHEET BƯỚC ĐỆM TRUNG GIAN ĐỂ ĐỐI SOÁT (Màu Xám Ghi)
                    l, q, a, _, _ = ghi_sheet_he_thong(mảng_bán_dương, "Du_Lieu_Ban_Duong", "595959", is_data_thô=True)
                    grand_summary_list.append(["Du_Lieu_Ban_Duong", "Thô dương độc lập", l, q, a, 0.0, a, "Nguyên bản đầu vào dương"])
                    
                    l, q, a, _, _ = ghi_sheet_he_thong(mảng_trả_âm, "Du_Lieu_Tra_Am", "595959", is_data_thô=True)
                    grand_summary_list.append(["Du_Lieu_Tra_Am", "Thô âm độc lập", l, q, a, 0.0, a, "Bảo toàn nguyên bản dấu âm"])

                    # B. GHI SHEET TỔNG HỢP THEO GIÁ BÌA (Màu Xanh Dương Đậm)
                    if bảng_gia_bia:
                        l, q, a, _, _ = ghi_sheet_he_thong(bảng_gia_bia, "TH_Hang_Ban_Gia_Bia", "002060")
                        grand_summary_list.append(["TH_Hang_Ban_Gia_Bia", "Gom mã (Giá bìa)", l, q, a, 0.0, a, "Chiết khấu bằng 0"])

                    # C. GHI SHEET TỔNG HỢP THEO CHIẾT KHẤU (Màu Xanh Lá Cây)
                    if bảng_chiet_khau:
                        l_ck_goc, q_ck_goc, a_ck_goc, v_ck_goc, s_ck_goc = ghi_sheet_he_thong(bảng_chiet_khau, "TH_Hang_Ban_Chiet_Khau", "339933", has_vat=True)
                        grand_summary_list.append(["TH_Hang_Ban_Chiet_Khau", "Gom mã + CK (Bán)", l_ck_goc, q_ck_goc, a_ck_goc, v_ck_goc, s_ck_goc, "Doanh thu xuất bán thực tế"])
                    else:
                        l_ck_goc, q_ck_goc, a_ck_goc, v_ck_goc, s_ck_goc = 0, 0.0, 0.0, 0.0, 0.0

                    # D. GHI SHEET TỔNG HỢP HÀNG TRẢ BẢO TOÀN DẤU ÂM (Màu Đỏ Đô)
                    if bảng_hang_tra:
                        l_tra_goc, q_tra_goc, a_tra_goc, v_tra_goc, s_tra_goc = ghi_sheet_he_thong(bảng_hang_tra, "Tong_Hop_Hang_Tra", "C00000", has_vat=True)
                        grand_summary_list.append(["Tong_Hop_Hang_Tra", "Gom mã + CK (Trả)", l_tra_goc, q_tra_goc, a_tra_goc, v_tra_goc, s_tra_goc, "Hàng trả giữ nguyên dấu âm"])
                    else:
                        l_tra_goc, q_tra_goc, a_tra_goc, v_tra_goc, s_tra_goc = 0, 0.0, 0.0, 0.0, 0.0

                    # E. CHẶT KHÚC 1000 DÒNG TÁCH HÓA ĐƠN TỪ TẬP CHIẾT KHẤU BÁN DƯƠNG (Màu Xanh Ngọc)
                    if bảng_chiet_khau:
                        batch_size = 1000
                        hd_count = 1
                        for start_pos in range(0, len(bảng_chiet_khau), batch_size):
                            batch_rows = bảng_chiet_khau[start_pos : start_pos + batch_size]
                            
                            # Đánh lại STT tịnh tiến từ 1 cho từng hóa đơn tách độc lập
                            data_hd_reindexed = []
                            for i_stt, r_data in enumerate(batch_rows, 1):
                                r_copy = list(r_data)
                                r_copy[0] = i_stt
                                data_hd_reindexed.append(r_copy)
                                
                            hd_name = f"HD {hd_count}"
                            l, q, a, vt, st_hd = ghi_sheet_he_thong(data_hd_reindexed, hd_name, "008080", has_vat=True)
                            grand_summary_list.append([hd_name, "Hóa đơn tách 1000 dòng", l, q, a, vt, st_hd, f"Tách đoạn từ dòng {start_pos+1}"])
                            hd_count += 1

                    # F. XÂY DỰNG SHEET MA TRẬN TĨNH ĐỐI SOÁT TỐI CAO `Tong_Ket_Chung` (Màu Đen Kế Toán)
                    columns_grand = ["Tên Sheet / Hạng mục", "Loại thuộc tính", "Số dòng", "Số lượng", "Trước Thuế", "VAT (5%)", "Sau Thuế", "Ghi chú"]
                    df_grand_final = pd.DataFrame(grand_summary_list, columns=columns_grand)
                    df_grand_final.to_excel(writer, index=False, sheet_name="Tong_Ket_Chung")
                    
                    ws_grand = writer.sheets["Tong_Ket_Chung"]
                    r_net = len(df_grand_final) + 2
                    
                    # CỘNG ĐẠI SỐ TRỰC TIẾP GIỮA BÁN (DƯƠNG) VÀ TRẢ (ÂM) - CHÍNH XÁC TUYỆT ĐỐI KHÔNG LỆCH DẤU
                    net_qty = q_ck_goc + q_tra_goc
                    net_amount = a_ck_goc + a_tra_goc
                    net_vat = v_ck_goc + v_tra_goc
                    net_after_tax = s_ck_goc + s_tra_goc
                    
                    ws_grand.cell(row=r_net, column=1).value = "DOANH THU THỰC TẾ ĐỐI SOÁT (BÁN + TRẢ)"
                    ws_grand.cell(row=r_net, column=2).value = "Cộng đại số PAB21"
                    ws_grand.cell(row=r_net, column=3).value = len(mảng_bán_dương) + len(mảng_trả_âm)
                    ws_grand.cell(row=r_net, column=4).value = net_qty
                    ws_grand.cell(row=r_net, column=5).value = net_amount
                    ws_grand.cell(row=r_net, column=6).value = net_vat
                    ws_grand.cell(row=r_net, column=7).value = net_after_tax
                    
                    trang_tri_sheet(ws_grand, "1F1F1F", total_row_type="grand")
                    
                    # Đẩy sheet Tong_Ket_Chung lên vị trí đầu tiên của Workbook
                    wb_obj = writer.book
                    sheets_order = [wb_obj.sheetnames[-1]] + wb_obj.sheetnames[:-1]
                    wb_obj._sheets = [wb_obj._sheets[wb_obj.sheetnames.index(name)] for name in sheets_order]

                st.success("🎉 PHƯƠNG PHÁP PAB21 ĐÃ XỬ LÝ HOÀN TẤT VÀ CHÍNH XÁC TUYỆT ĐỐI!")
                
                # Hiển thị nhanh bảng đối soát lên giao diện Streamlit
                t_total, t_ban_dem, t_tra_dem = st.tabs(["📋 Bảng Tổng Kết Đối Soát", "📥 Mảng Đệm Dương", "📤 Mảng Đệm Âm (Bảo Toàn Dấu)"])
                with t_total:
                    st.dataframe(df_grand_final, use_container_width=True)
                    st.metric(label="Doanh Thu Thuần Thực Tế Cấu Trừ (Sau Thuế)", value=f"{net_after_tax:,.2f} đ")
                with t_ban_dem:
                    st.dataframe(pd.DataFrame(mảng_bán_dương, columns=columns_standard), use_container_width=True)
                with t_tra_dem:
                    st.dataframe(pd.DataFrame(mảng_trả_âm, columns=columns_standard), use_container_width=True)

                st.download_button(
                    label="📥 TẢI FILE MASTER REPORT QUY CHUẨN PAB21",
                    data=out_report.getvalue(),
                    file_name=f"Master_Report_PAB21_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Lỗi nghiêm trọng phát sinh tại Giai đoạn 2 (PAB21): {str(e)}")
