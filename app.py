import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =============================================================================
# CẤU HÌNH HỆ THỐNG & GIAO DIỆN
# =============================================================================
st.set_page_config(
    page_title="Hệ thống Quản lý Tiêu chí Chất lượng Vận hành Tín dụng",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Hệ thống Xử lý & Phân tích Dữ liệu Tín dụng")
st.markdown("---")

# Tạo Sidebar điều hướng giữa 4 tiêu chí
menu = st.sidebar.radio(
    "CHỌN TIÊU CHÍ XỬ LÝ DỮ LIỆU",
    [
        "Tiêu chí 1: Dư nợ cơ cấu",
        "Tiêu chí 2: Miễn giảm lãi",
        "Tiêu chí 3: Lũy kế nợ quá hạn",
        "Tiêu chí 4: Tín dụng thanh khoản thấp"
    ]
)

# -----------------------------------------------------------------------------
# HÀM ĐỌC FILE THÔNG MINH (CHẤP NHẬN XLSX & XLS, TRÁNH LỖI STRUCT.ERROR / HTML)
# -----------------------------------------------------------------------------
def read_excel_smart(uploaded_files, target_month=None):
    if not uploaded_files:
        return pd.DataFrame()
    
    if not isinstance(uploaded_files, list):
        uploaded_files = [uploaded_files]
        
    list_df = []
    for f in uploaded_files:
        f_bytes = f.read()
        f.seek(0)  # Reset con trỏ file về đầu
        df_temp = pd.DataFrame()
        
        try:
            # Bước 1: Thử đọc theo engine tương ứng dựa vào đuôi file
            engine_pick = "xlrd" if f.name.endswith('.xls') else "openpyxl"
            df_temp = pd.read_excel(io.BytesIO(f_bytes), engine=engine_pick, dtype=str)
        except Exception:
            try:
                # Bước 2: Thử ép chéo cấu trúc engine đề phòng file sai đuôi mở rộng
                engine_pick = "openpyxl" if f.name.endswith('.xls') else "xlrd"
                df_temp = pd.read_excel(io.BytesIO(f_bytes), engine=engine_pick, dtype=str)
            except Exception:
                try:
                    # Bước 3: Thử đọc trực tiếp định dạng HTML (đối với file Core banking giả lập đuôi .xls)
                    df_list = pd.read_html(io.BytesIO(f_bytes))
                    df_temp = df_list[0].astype(str) if df_list else pd.DataFrame()
                except Exception as e:
                    st.error(f"❌ Không thể đọc được file: {f.name}. Lỗi chi tiết: {str(e)}")
                    continue
                    
        if not df_temp.empty:
            df_temp["TEN_FILE"] = f.name
            if target_month:
                df_temp["THANG"] = target_month
            list_df.append(df_temp)
            
    if list_df:
        return pd.concat(list_df, ignore_index=True)
    return pd.DataFrame()

# Hàm bổ trợ để chuyển đổi DataFrame sang file Excel phục vụ nút Tải về
def to_excel_download(df, sheet_name="Sheet1"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


# ==========================================
# TIÊU CHÍ 1: DƯ NỢ CƠ CẤU
# ==========================================
if menu == "Tiêu chí 1: Dư nợ cơ cấu":
    st.header("📌 Tiêu chí 1: Đối chiếu & Phân bổ Dư nợ cơ cấu")
    st.info("💡 Điểm cải tiến: Hệ thống tự động chấp nhận cả file dạng .xls và .xlsx.")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("🗓️ Dữ liệu Tháng này (T)")
        files_crm32_nay = st.file_uploader("Upload các file CRM32 Tháng này (.xls, .xlsx)", accept_multiple_files=True, key="t1_m32_n")
        files_crm4_nay = st.file_uploader("Upload các file CRM4 Tháng này (.xls, .xlsx)", accept_multiple_files=True, key="t1_m4_n")
    
    with col2:
        st.subheader("🗓️ Dữ liệu Tháng trước (T-1)")
        files_crm32_truoc = st.file_uploader("Upload các file CRM32 Tháng trước (.xls, .xlsx)", accept_multiple_files=True, key="t1_m32_t")
        files_crm4_truoc = st.file_uploader("Upload các file CRM4 Tháng trước (.xls, .xlsx)", accept_multiple_files=True, key="t1_m4_t")

    st.subheader("📂 Dữ liệu Phụ trợ")
    file_sddp = st.file_uploader("Upload file Số dư dự phòng DPRR_T5 (.xls, .xlsx)", key="t1_sddp")

    if st.button("🚀 Chạy Xử lý Tiêu chí 1"):
        if files_crm32_nay and files_crm4_nay and files_crm32_truoc and files_crm4_truoc and file_sddp:
            with st.spinner("Đang tính toán dư nợ cơ cấu..."):
                # Đọc cấu trúc Tháng này bằng hàm smart reader
                df_crm32_nay = read_excel_smart(files_crm32_nay)
                df_crm4_nay = read_excel_smart(files_crm4_nay)
                
                # Đọc cấu trúc Tháng trước bằng hàm smart reader
                df_crm32_truoc = read_excel_smart(files_crm32_truoc)
                df_crm4_truoc = read_excel_smart(files_crm4_truoc)

                # Chuẩn hóa viết hoa toàn bộ tên cột để tránh lỗi lệch font hệ thống
                for _df in [df_crm32_nay, df_crm4_nay, df_crm32_truoc, df_crm4_truoc]:
                    _df.columns = _df.columns.str.strip().str.upper()

                # Xử lý tháng này
                df_crm32_cc_nay = df_crm32_nay[df_crm32_nay["SCHM_DESC"].astype(str).str.upper().str.replace(r"\s+", " ", regex=True).str.contains("CO CAU", na=False)].drop_duplicates(subset=["CUSTSEQLN"])
                df_join_nay = df_crm32_cc_nay.merge(df_crm4_nay, left_on="CUSTSEQLN", right_on="CIF_KH_VAY", how="inner")
                df_join_nay["DU_NO_PHAN_BO_QUY_DOI"] = pd.to_numeric(df_join_nay["DU_NO_PHAN_BO_QUY_DOI"], errors="coerce").fillna(0)
                
                df_kq_nay = df_join_nay.groupby(["CIF_KH_VAY", "SCHM_DESC","NHOM_NO"], as_index=False)["DU_NO_PHAN_BO_QUY_DOI"].sum().rename(columns={"DU_NO_PHAN_BO_QUY_DOI": "TONG_DU_NO_THANG_5","NHOM_NO":"NHOM_NO_T5"})
                df_kh = df_join_nay[["CIF_KH_VAY", "KHACH_HANG"]].drop_duplicates("CIF_KH_VAY")
                df_kq_nay = df_kq_nay.merge(df_kh, on="CIF_KH_VAY", how="left")

                # Xử lý tháng trước
                df_crm32_cc_truoc = df_crm32_truoc[df_crm32_truoc["SCHM_DESC"].astype(str).str.upper().str.replace(r"\s+", " ", regex=True).str.contains("CO CAU", na=False)].drop_duplicates(subset=["CUSTSEQLN"])
                df_join_truoc = df_crm32_cc_truoc.merge(df_crm4_truoc, left_on="CUSTSEQLN", right_on="CIF_KH_VAY", how="inner")
                df_join_truoc["DU_NO_PHAN_BO_QUY_DOI"] = pd.to_numeric(df_join_truoc["DU_NO_PHAN_BO_QUY_DOI"], errors="coerce").fillna(0)
                
                df_kq_truoc = df_join_truoc.groupby(["CIF_KH_VAY", "SCHM_DESC", "NHOM_NO"], as_index=False)["DU_NO_PHAN_BO_QUY_DOI"].sum().rename(columns={"DU_NO_PHAN_BO_QUY_DOI": "TONG_DU_NO_THANG_4","NHOM_NO":"NHOM_NO_T4"})
                df_kh_trc = df_join_truoc[["CIF_KH_VAY", "KHACH_HANG"]].drop_duplicates("CIF_KH_VAY")
                df_kq_truoc = df_kq_truoc.merge(df_kh_trc, on="CIF_KH_VAY", how="left")

                # Đối chiếu kết quả
                df_kq_nay_rename = df_kq_nay.rename(columns={"CIF_KH_VAY": "CIF_KH_VAY_NAY","SCHM_DESC": "SCHM_DESC_NAY", "KHACH_HANG":"TEN_KH_T5"})
                df_kq_truoc_rename = df_kq_truoc.rename(columns={"CIF_KH_VAY": "CIF_KH_VAY_TRUOC","SCHM_DESC": "SCHM_DESC_TRUOC", "KHACH_HANG":"TEN_KH_T4"}
                )
                
                df_doi_chieu = df_kq_nay_rename.merge(df_kq_truoc_rename, left_on=["CIF_KH_VAY_NAY", "SCHM_DESC_NAY"], right_on=["CIF_KH_VAY_TRUOC", "SCHM_DESC_TRUOC"], how="outer").fillna(0)
                df_doi_chieu["CHENH_LECH"] = df_doi_chieu["TONG_DU_NO_THANG_5"] - df_doi_chieu["TONG_DU_NO_THANG_4"]
                df_doi_chieu = df_doi_chieu.sort_values("CHENH_LECH", ascending=False)

                # Đọc Map số dư dự phòng dùng Smart Reader
                df_sddp = read_excel_smart(file_sddp)
                df_sddp.columns = df_sddp.columns.str.strip().str.upper()
                df_sddp["PHAT_SINH_NO"] = pd.to_numeric(df_sddp["PHAT_SINH_NO"], errors="coerce").fillna(0)
                
                df_doi_chieu["CIF_KH_VAY_NAY"] = df_doi_chieu["CIF_KH_VAY_NAY"].astype(str).str.strip().str.replace(".0","", regex=False)
                df_sddp["CIF"] = df_sddp["CIF"].astype(str).str.strip().str.replace(".0","", regex=False)
                
                df_sddp_sum = df_sddp.groupby("CIF")["PHAT_SINH_NO"].sum()
                df_doi_chieu["DPRR"] = df_doi_chieu["CIF_KH_VAY_NAY"].map(df_sddp_sum)

                st.success("🎉 Xử lý thành công!")
                st.dataframe(df_doi_chieu)

                # Cho phép người dùng tải xuống kết quả dưới định dạng excel
                excel_data = to_excel_download(df_doi_chieu, sheet_name="DU_NO_CO_CAU")
                st.download_button("📥 Tải File Kết Quả Excel", data=excel_data, file_name="DU_NO_CO_CAU_OUTPUT.xlsx")
        else:
            st.error("⚠️ Vui lòng cung cấp đầy đủ các file đầu vào bắt buộc!")

# ==========================================
# TIÊU CHÍ 2: MIỄN GIẢM LÃI
# ==========================================
elif menu == "Tiêu chí 2: Miễn giảm lãi":
    st.header("📌 Tiêu chí 2: Tổng hợp dữ liệu Miễn giảm lãi")
    
    col1, col2 = st.columns(2)
    with col1:
        file_hlawint = st.file_uploader("Upload file Màn hình HLAWINT (.xls, .xlsx)", key="t2_hl")
        file_noibang = st.file_uploader("Upload file Thu nợ Nội bảng (.xls, .xlsx)", key="t2_nb")
    with col2:
        file_ngoaibang = st.file_uploader("Upload file Thu nợ Ngoại bảng (.xls, .xlsx)", key="t2_ngb")
        file_thulai = st.file_uploader("Upload file Thu lãi (.xls, .xlsx)", key="t2_tl")

    if st.button("🚀 Chạy Xử lý Tiêu chí 2"):
        if file_hlawint and file_noibang and file_ngoaibang and file_thulai:
            with st.spinner("Đang tổng hợp dữ liệu..."):
                # 1. Khởi tạo dữ liệu từ HLAWINT dùng Smart Reader
                df = read_excel_smart(file_hlawint)
                cols = ['SOL_ID','SOL_DESC','CIF_ID','CUST_NAME','INTAMT_VND','LY_DO']
                df = df[cols]
                
                df_unique = df.drop_duplicates(subset='CIF_ID')
                df['INTAMT_VND'] = pd.to_numeric(df['INTAMT_VND'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                sum_intamt = df.groupby('CIF_ID', as_index=False)['INTAMT_VND'].sum()
                result = pd.merge(sum_intamt, df_unique[['CIF_ID','SOL_ID','SOL_DESC','CUST_NAME','LY_DO']], on='CIF_ID', how='left')

                # 2. Map file ngoại bảng dùng Smart Reader
                dprr = read_excel_smart(file_ngoaibang)
                dprr['Số tiền thu'] = pd.to_numeric(dprr['Số tiền thu'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                dprr_sum = dprr.groupby('CIF 9 SO', as_index=False)['Số tiền thu'].sum().rename(columns={'Số tiền thu':'Thu_ngoai_bang', 'CIF 9 SO':'CIF_ID'})
                result = result.merge(dprr_sum, on='CIF_ID', how='left')

                # 3. Map file nội bảng (Thu gốc) dùng Smart Reader
                thu_goc = read_excel_smart(file_noibang)
                thu_goc['Số tiền thu'] = pd.to_numeric(thu_goc['Số tiền thu'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                thu_goc_sum = thu_goc.groupby('CIF 9 SO', as_index=False)['Số tiền thu'].sum().rename(columns={'Số tiền thu':'Thu_goc_trong_thang', 'CIF 9 SO':'CIF_ID'})
                result = result.merge(thu_goc_sum, on='CIF_ID', how='left')

                # 4. Map file thu lãi dùng Smart Reader
                thu_lai = read_excel_smart(file_thulai)
                thu_lai['Số tiền thu'] = pd.to_numeric(thu_lai['Số tiền thu'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                thu_lai_sum = thu_lai.groupby('CIF 9 SO', as_index=False)['Số tiền thu'].sum().rename(columns={'Số tiền thu':'Thu_lai_trong_thang', 'CIF 9 SO':'CIF_ID'})
                result = result.merge(thu_lai_sum, on='CIF_ID', how='left')

                # Tính tổng thu và sinh cột theo form báo cáo định sẵn
                result['Tong_thu'] = result['Thu_goc_trong_thang'].fillna(0) + result['Thu_lai_trong_thang'].fillna(0) + result['Thu_ngoai_bang'].fillna(0)
                
                new_cols = ['Thu_thieu_lai_trong_han', 'Mien_goc', 'Mien_lai_trong_han', 'Mien_lai_qua_han', 'Mien_lai_phat', 'Ghi_chu', 'Con_ton', 'Ty_quan', 'So_thu_phi', 'note']
                for col in new_cols:
                    result[col] = 0

                result['Tham_quyen_phe_duyet'] = result['LY_DO']
                muc6 = df.groupby('CIF_ID')['INTAMT_VND'].sum().reset_index().rename(columns={'INTAMT_VND': 'muc6'})
                result = result.merge(muc6, on='CIF_ID', how='left')
                result['check_tongMGL'] = (result['Mien_goc'] + result['Mien_lai_trong_han'] + result['Mien_lai_qua_han'] + result['Mien_lai_phat']) - result['muc6']

                # Sắp xếp đúng cấu trúc cột đầu ra mẫu
                order_cols = ['SOL_ID','SOL_DESC','CIF_ID','CUST_NAME','Tong_thu','Thu_ngoai_bang','Thu_goc_trong_thang','Thu_lai_trong_thang','Thu_thieu_lai_trong_han','Mien_goc','Mien_lai_trong_han','Mien_lai_qua_han','Mien_lai_phat','Tham_quyen_phe_duyet','Ghi_chu','Con_ton','Ty_quan','So_thu_phi','muc6','check_tongMGL','note']
                result = result[order_cols]

                st.success("🎉 Tạo bảng Miễn giảm lãi thành công!")
                st.dataframe(result)
                
                excel_data = to_excel_download(result, sheet_name="XulyNo")
                st.download_button("📥 Tải Báo Cáo Xử Lý Nợ (.xlsx)", data=excel_data, file_name="XulyNo_T5_Output.xlsx")
        else:
            st.error("⚠️ Vui lòng tải lên đầy đủ các file dữ liệu phục vụ tính toán!")

# ==========================================
# TIÊU CHÍ 3: LŨY KẾ NỢ QUÁ HẠN
# ==========================================
elif menu == "Tiêu chí 3: Lũy kế nợ quá hạn":
    st.header("📌 Tiêu chí 3: Theo dõi biến động Lũy kế Nợ quá hạn qua các tháng")
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("Cấu hình chuỗi thời gian")
    thang_bat_dau = st.sidebar.text_input("Từ tháng (YYYY-MM)", "2026-01")
    thang_ket_thuc = st.sidebar.text_input("Đến tháng (YYYY-MM)", "2026-05")

    # Sinh danh sách các tháng động
    ds_thang = pd.period_range(start=thang_bat_dau, end=thang_ket_thuc, freq="M").astype(str).tolist()
    
    st.markdown(f"**Yêu cầu nạp file dữ liệu CRM32 cho từng tháng trong chuỗi:** `{ds_thang}`")
    
    uploaded_months_data = {}
    for m in ds_thang:
        uploaded_months_data[m] = st.file_uploader(f"Tháng {m}: Nạp file CRM32 (.xls, .xlsx)", accept_multiple_files=True, key=f"t3_{m}")

    if st.button("🚀 Chạy Tính Toán Lũy Kế"):
        all_month_data = []
        valid = True
        
        for m in ds_thang:
            if not uploaded_months_data[m]:
                st.warning(f"Chưa có tệp dữ liệu cho tháng {m}!")
                valid = False
            else:
                # Đọc tích hợp nhiều file của từng tháng bằng Smart Reader
                df_m_temp = read_excel_smart(uploaded_months_data[m], target_month=m)
                if not df_m_temp.empty:
                    all_month_data.append(df_m_temp)
                    
        if valid and len(all_month_data) > 0:
            with st.spinner("Hệ thống đang tổng hợp dữ liệu chuỗi thời gian..."):
                df = pd.concat(all_month_data, ignore_index=True)
                df.columns = df.columns.str.strip().str.upper()
                
                col_brcd, col_chi_nhanh, col_du_no, col_nhom_no, col_ngay_giai_ngan = "BRCD", "CHI_NHANH", "DU_NO_QUY_DOI", "NHOM_NO_THEO_CIF", "NGAY_GIAI_NGAN"
                
                df[col_brcd] = df[col_brcd].astype(str).str.strip()
                df[col_chi_nhanh] = df[col_chi_nhanh].astype(str).str.strip()
                df[col_du_no] = pd.to_numeric(df[col_du_no], errors="coerce").fillna(0)
                df[col_nhom_no] = pd.to_numeric(df[col_nhom_no], errors="coerce")
                df[col_ngay_giai_ngan] = pd.to_datetime(df[col_ngay_giai_ngan], errors="coerce", dayfirst=True)

                ds_bao_cao_thang = []
                for m in ds_thang:
                    df_thang = df[df["THANG"] == m].copy()
                    ngay_bao_cao = pd.to_datetime(m + "-01") + pd.offsets.MonthEnd(0)
                    df_thang["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] = (ngay_bao_cao - df_thang[col_ngay_giai_ngan]).dt.days
                    df_qh = df_thang[df_thang[col_nhom_no].isin([2, 3, 4, 5])].copy()

                    giai_ngan_3_thang_qh = df_qh[(df_qh["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] >= 0) & (df_qh["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] <= 90)].groupby([col_brcd, col_chi_nhanh], as_index=False).agg(GIAI_NGAN_3_THANG_QH=(col_du_no, "sum"))
                    giai_ngan_6_thang_qh = df_qh[(df_qh["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] >= 0) & (df_qh["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] <= 180)].groupby([col_brcd, col_chi_nhanh], as_index=False).agg(GIAI_NGAN_6_THANG_QH=(col_du_no, "sum"))
                    base = df_thang.groupby([col_brcd, col_chi_nhanh], as_index=False).agg(TONG_DU_NO_THANG=(col_du_no, "sum"))

                    res_m = base.merge(giai_ngan_3_thang_qh, on=[col_brcd, col_chi_nhanh], how="left").merge(giai_ngan_6_thang_qh, on=[col_brcd, col_chi_nhanh], how="left").fillna(0)
                    res_m["THANG"] = m
                    res_m["NGAY_BAO_CAO"] = ngay_bao_cao
                    ds_bao_cao_thang.append(res_m)

                bao_cao_all_month = pd.concat(ds_bao_cao_thang, ignore_index=True)
                bao_cao_all_month["THANG_DATE"] = pd.to_datetime(bao_cao_all_month["THANG"] + "-01")
                bao_cao_all_month = bao_cao_all_month.sort_values(by=[col_brcd, "THANG_DATE"])

                # Cấp số cộng lũy kế tịnh tiến (Đã xử lý bỏ nháy kép thừa)
                bao_cao_all_month["LUY_KE_GIAI_NGAN_3_THANG_QH"] = bao_cao_all_month.groupby([col_brcd, col_chi_nhanh])["GIAI_NGAN_3_THANG_QH"].cumsum()
                bao_cao_all_month["LUY_KE_GIAI_NGAN_6_THANG_QH"] = bao_cao_all_month.groupby([col_brcd, col_chi_nhanh])["GIAI_NGAN_6_THANG_QH"].cumsum()

                # Đổi trục Pivot làm báo cáo ngang gọn đẹp
                pivot_luy_ke = bao_cao_all_month.pivot_table(index=[col_brcd, col_chi_nhanh], columns="THANG", values=["GIAI_NGAN_3_THANG_QH", "GIAI_NGAN_6_THANG_QH", "LUY_KE_GIAI_NGAN_3_THANG_QH", "LUY_KE_GIAI_NGAN_6_THANG_QH"], aggfunc="sum")
                pivot_luy_ke.columns = [f"{ct}_{th}" for ct, th in pivot_luy_ke.columns]
                pivot_luy_ke = pivot_luy_ke.reset_index()

                st.success("📊 Bảng dữ liệu biến động ngang thu được:")
                st.dataframe(pivot_luy_ke)

                # Xuất Workbook đa dạng Sheet cho khách hàng download
                out_bytes = io.BytesIO()
                with pd.ExcelWriter(out_bytes, engine="openpyxl") as wr:
                    bao_cao_all_month.to_excel(wr, sheet_name="Data luy ke theo thang", index=False)
                    pivot_luy_ke.to_excel(wr, sheet_name="Bang ngang luy ke", index=False)
                
                st.download_button("📥 Tải File Lũy Kế Đa Bản Tấm (.xlsx)", data=out_bytes.getvalue(), file_name=f"Luy_Ke_Qua_Han_{thang_bat_dau}_To_{thang_ket_thuc}.xlsx")

# ==========================================
# TIÊU CHÍ 4: THANH KHOẢN THẤP & AUTO STYLING EXCEL
# ==========================================
elif menu == "Tiêu chí 4: Tín dụng thanh khoản thấp":
    st.header("📌 Tiêu chí 4: So sánh & Tự động tô màu Báo cáo Thanh khoản thấp")
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("Thiết lập tháng đánh giá")
    thang_t_input = st.sidebar.text_input("Nhập tháng T (YYYY-MM)", "2026-04")

    # Tự quy đổi tháng liên kề liền trước
    thang_t_date = pd.to_datetime(thang_t_input + "-01")
    thang_t_1_date = thang_t_date - pd.DateOffset(months=1)
    thang_t = thang_t_date.strftime("%Y-%m")
    thang_t_1 = thang_t_1_date.strftime("%Y-%m")

    st.write(f"🔄 **Hệ thống tự động thiết lập cặp so sánh:** Tháng T: **{thang_t}** với Tháng T-1: **{thang_t_1}**")

    col1, col2 = st.columns(2)
    with col1:
        files_tc4_t = st.file_uploader(f"Upload các file CRM32 của Tháng T ({thang_t})", accept_multiple_files=True, key="t4_t")
    with col2:
        files_tc4_t1 = st.file_uploader(f"Upload các file CRM32 của Tháng T-1 ({thang_t_1})", accept_multiple_files=True, key="t4_t1")

    if st.button("🚀 Chạy Đối Chiếu & Thiết Kế Form Đẹp"):
        if files_tc4_t and files_tc4_t1:
            with st.spinner("Đang biên tập & thiết kế format bảng Excel tự động..."):
                # Đọc xử lý bằng Smart Reader chấp nhận cả .xls/.xlsx
                df_t = read_excel_smart(files_tc4_t, target_month=thang_t)
                df_t_1 = read_excel_smart(files_tc4_t1, target_month=thang_t_1)

                df = pd.concat([df_t_1, df_t], ignore_index=True)
                df.columns = df.columns.str.strip().str.upper()

                col_brcd, col_chi_nhanh, col_du_no, col_nhom_no, col_ngay_giai_ngan = "BRCD", "CHI_NHANH", "DU_NO_QUY_DOI", "NHOM_NO_THEO_CIF", "NGAY_GIAI_NGAN"
                df[col_brcd] = df[col_brcd].astype(str).str.strip()
                df[col_chi_nhanh] = df[col_chi_nhanh].astype(str).str.strip()
                df[col_du_no] = pd.to_numeric(df[col_du_no], errors="coerce").fillna(0)
                df[col_nhom_no] = pd.to_numeric(df[col_nhom_no], errors="coerce")
                df[col_ngay_giai_ngan] = pd.to_datetime(df[col_ngay_giai_ngan], errors="coerce")

                def process_one_month(data, target_m):
                    df_m = data[data["THANG"] == target_m].copy()
                    ngay_bc = pd.to_datetime(target_m + "-01") + pd.offsets.MonthEnd(0)
                    df_m["SO_NGAY"] = (ngay_bc - df_m[col_ngay_giai_ngan]).dt.days

                    du_no = df_m.groupby([col_brcd, col_chi_nhanh], as_index=False).agg(DU_NO=(col_du_no, "sum"))
                    du_no_qh = df_m[df_m[col_nhom_no].isin([2, 3, 4, 5])].groupby([col_brcd, col_chi_nhanh], as_index=False).agg(DU_NO_QH=(col_du_no, "sum"))
                    du_no_xau = df_m[df_m[col_nhom_no].isin([3, 4, 5])].groupby([col_brcd, col_chi_nhanh], as_index=False).agg(DU_NO_XAU=(col_du_no, "sum"))
                    
                    g3 = df_m[(df_m["SO_NGAY"] >= 0) & (df_m["SO_NGAY"] <= 90) & (df_m[col_nhom_no].isin([2, 3, 4, 5]))].groupby([col_brcd, col_chi_nhanh], as_index=False).agg(GIAI_NGAN_3_THANG_QH=(col_du_no, "sum"))
                    g6 = df_m[(df_m["SO_NGAY"] >= 0) & (df_m["SO_NGAY"] <= 180) & (df_m[col_nhom_no].isin([2, 3, 4, 5]))].groupby([col_brcd, col_chi_nhanh], as_index=False).agg(GIAI_NGAN_6_THANG_QH=(col_du_no, "sum"))

                    res = du_no.merge(du_no_qh, on=[col_brcd, col_chi_nhanh], how="left").merge(du_no_xau, on=[col_brcd, col_chi_nhanh], how="left").merge(g3, on=[col_brcd, col_chi_nhanh], how="left").merge(g6, on=[col_brcd, col_chi_nhanh], how="left").fillna(0)
                    res["TY_LE_QH"] = np.where(res["DU_NO"] != 0, res["DU_NO_QH"] / res["DU_NO"], 0)
                    res["TY_LE_NPL"] = np.where(res["DU_NO"] != 0, res["DU_NO_XAU"] / res["DU_NO"], 0)
                    return res

                bc_t_1_data = process_one_month(df, thang_t_1)
                bc_t_data = process_one_month(df, thang_t)

                # Rename gộp bảng
                bc_t_1_rename = bc_t_1_data.rename(columns={"DU_NO": "DU_NO_T_1", "DU_NO_QH": "DU_NO_QH_T_1", "DU_NO_XAU": "DU_NO_XAU_T_1", "TY_LE_QH": "TY_LE_QH_T_1", "TY_LE_NPL": "TY_LE_NPL_T_1", "GIAI_NGAN_3_THANG_QH": "GIAI_NGAN_3_THANG_QH_T_1", "GIAI_NGAN_6_THANG_QH": "GIAI_NGAN_6_THANG_QH_T_1"})
                bc_t_rename = bc_t_data.rename(columns={"DU_NO": "DU_NO_T", "DU_NO_QH": "DU_NO_QH_T", "DU_NO_XAU": "DU_NO_XAU_T", "TY_LE_QH": "TY_LE_QH_T", "TY_LE_NPL": "TY_LE_NPL_T", "GIAI_NGAN_3_THANG_QH": "GIAI_NGAN_3_THANG_QH_T", "GIAI_NGAN_6_THANG_QH": "GIAI_NGAN_6_THANG_QH_T"})

                compare = bc_t_1_rename[[col_brcd, col_chi_nhanh, "DU_NO_T_1", "DU_NO_QH_T_1", "DU_NO_XAU_T_1", "TY_LE_QH_T_1", "TY_LE_NPL_T_1", "GIAI_NGAN_3_THANG_QH_T_1", "GIAI_NGAN_6_THANG_QH_T_1"]].merge(bc_t_rename[[col_brcd, col_chi_nhanh, "DU_NO_T", "DU_NO_QH_T", "DU_NO_XAU_T", "TY_LE_QH_T", "TY_LE_NPL_T", "GIAI_NGAN_3_THANG_QH_T", "GIAI_NGAN_6_THANG_QH_T"]], on=[col_brcd, col_chi_nhanh], how="outer").fillna(0)
                
                compare["SS_DU_NO"] = compare["DU_NO_T"] - compare["DU_NO_T_1"]
                compare["SS_QH"] = compare["DU_NO_QH_T"] - compare["DU_NO_QH_T_1"]
                compare["SS_NPL"] = compare["DU_NO_XAU_T"] - compare["DU_NO_XAU_T_1"]
                compare["SS_TY_LE_QH"] = compare["TY_LE_QH_T"] - compare["TY_LE_QH_T_1"]
                compare["SS_TY_LE_NPL"] = compare["TY_LE_NPL_T"] - compare["TY_LE_NPL_T_1"]
                compare["SS_GIAI_NGAN_3_THANG_QH"] = compare["GIAI_NGAN_3_THANG_QH_T"] - compare["GIAI_NGAN_3_THANG_QH_T_1"]
                compare["SS_GIAI_NGAN_6_THANG_QH"] = compare["GIAI_NGAN_6_THANG_QH_T"] - compare["GIAI_NGAN_6_THANG_QH_T_1"]

                # Cấu hình lại tên hiển thị đẹp mắt
                final_df = compare.rename(columns={
                    col_brcd: "BRCD", col_chi_nhanh: "CHI_NHANH",
                    "DU_NO_T_1": "Dư nợ_T-1", "DU_NO_QH_T_1": "Dư nợ QH_T-1", "DU_NO_XAU_T_1": "Dư nợ xấu_T-1", "TY_LE_QH_T_1": "%QH_T-1", "TY_LE_NPL_T_1": "%NPL_T-1", "GIAI_NGAN_3_THANG_QH_T_1": "Giải ngân 3 tháng quá hạn_T-1", "GIAI_NGAN_6_THANG_QH_T_1": "Giải ngân 6 tháng quá hạn_T-1",
                    "DU_NO_T": "Dư nợ_T", "DU_NO_QH_T": "Dư nợ QH_T", "DU_NO_XAU_T": "Dư nợ xấu_T", "TY_LE_QH_T": "%QH_T", "TY_LE_NPL_T": "%NPL_T", "GIAI_NGAN_3_THANG_QH_T": "Giải ngân 3 tháng quá hạn_T", "GIAI_NGAN_6_THANG_QH_T": "Giải ngân 6 tháng quá hạn_T",
                    "SS_DU_NO": "Dư nợ", "SS_QH": "QH", "SS_NPL": "NPL", "SS_TY_LE_QH": "%QH", "SS_TY_LE_NPL": "%NPL", "SS_GIAI_NGAN_3_THANG_QH": "Giải ngân 3 tháng quá hạn", "SS_GIAI_NGAN_6_THANG_QH": "Giải ngân 6 tháng quá hạn"
                })

                # Chèn dòng TOTAL lên đầu
                total_row = {"BRCD": "TOTAL", "CHI_NHANH": "Tổng cộng"}
                money_cols = ["Dư nợ_T-1", "Dư nợ QH_T-1", "Dư nợ xấu_T-1", "Giải ngân 3 tháng quá hạn_T-1", "Giải ngân 6 tháng quá hạn_T-1", "Dư nợ_T", "Dư nợ QH_T", "Dư nợ xấu_T", "Giải ngân 3 tháng quá hạn_T", "Giải ngân 6 tháng quá hạn_T", "Dư nợ", "QH", "NPL", "Giải ngân 3 tháng quá hạn", "Giải ngân 6 tháng quá hạn"]
                
                for col in money_cols:
                    total_row[col] = final_df[col].sum()

                total_row["%QH_T-1"] = total_row["Dư nợ QH_T-1"] / total_row["Dư nợ_T-1"] if total_row["Dư nợ_T-1"] != 0 else 0
                total_row["%NPL_T-1"] = total_row["Dư nợ xấu_T-1"] / total_row["Dư nợ_T-1"] if total_row["Dư nợ_T-1"] != 0 else 0
                total_row["%QH_T"] = total_row["Dư nợ QH_T"] / total_row["Dư nợ_T"] if total_row["Dư nợ_T"] != 0 else 0
                total_row["%NPL_T"] = total_row["Dư nợ xấu_T"] / total_row["Dư nợ_T"] if total_row["Dư nợ_T"] != 0 else 0
                total_row["%QH"] = total_row["%QH_T"] - total_row["%QH_T-1"]
                total_row["%NPL"] = total_row["%NPL_T"] - total_row["%NPL_T-1"]

                final_df = pd.concat([pd.DataFrame([total_row]), final_df], ignore_index=True)
                
                st.write("📊 Xem trước kết quả xử lý đối chiếu:")
                st.dataframe(final_df.head(10))

                # Ghi dữ liệu vào openpyxl và thực hiện format đồ họa màu sắc như mã cũ
                out_styled = io.BytesIO()
                with pd.ExcelWriter(out_styled, engine="openpyxl") as writer:
                    final_df.to_excel(writer, sheet_name="Tong hop", index=False, startrow=3)
                    bc_t_1_data.to_excel(writer, sheet_name=f"Data {thang_t_1}", index=False)
                    bc_t_data.to_excel(writer, sheet_name=f"Data {thang_t}", index=False)

                # Load lại bộ đệm openpyxl để vẽ mỹ thuật
                wb = load_workbook(out_styled)
                ws = wb["Tong hop"]
                max_row, max_col = ws.max_row, ws.max_column

                # Thiết lập bảng màu chuẩn mực của ngân hàng
                fill_yellow = PatternFill("solid", fgColor="FFFF00")
                fill_green = PatternFill("solid", fgColor="C6E0B4")
                fill_gray = PatternFill("solid", fgColor="D9D9D9")
                fill_total = PatternFill("solid", fgColor="A9D08E")
                font_red = Font(color="FF0000", bold=True)
                font_bold = Font(bold=True)
                align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
                align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
                thin = Side(border_style="thin", color="000000")
                border = Border(left=thin, right=thin, top=thin, bottom=thin)

                # Merge cell tiêu đề nhóm lớn cấp cao
                ws.merge_cells("A1:B1"); ws.cell(1,1).value = ""
                ws.merge_cells("C1:I1"); ws.cell(1,3).value = f"Tháng T-1 ({thang_t_1})"
                ws.merge_cells("J1:P1"); ws.cell(1,10).value = f"Tháng T ({thang_t})"
                ws.merge_cells("Q1:W1"); ws.cell(1,17).value = "So sánh (T) - (T-1)"

                for col in range(3, 24): ws.cell(row=2, column=col).value = col - 2

                for row in [1, 2, 4]:
                    for col in range(1, max_col + 1):
                        cell = ws.cell(row=row, column=col)
                        cell.alignment = align_center; cell.border = border; cell.font = font_bold

                for col in range(3, 17):
                    ws.cell(1, col).fill = fill_yellow; ws.cell(2, col).fill = fill_yellow; ws.cell(4, col).fill = fill_yellow
                for col in range(17, 24):
                    ws.cell(1, col).fill = fill_green; ws.cell(2, col).fill = fill_green; ws.cell(4, col).fill = fill_green
                for col in range(1, 3): ws.cell(4, col).fill = fill_gray

                ws.cell(1, 3).font = font_red; ws.cell(1, 10).font = font_red; ws.cell(1, 17).font = font_red

                # Định dạng dữ liệu: Số và Tỷ lệ %
                percent_cols = ["%QH_T-1", "%NPL_T-1", "%QH_T", "%NPL_T", "%QH", "%NPL"]
                for row in range(1, max_row + 1):
                    for col in range(1, max_col + 1):
                        cell = ws.cell(row=row, column=col)
                        cell.border = border
                        if row >= 5: cell.alignment = align_center

                for row in range(5, max_row + 1): ws.cell(row=row, column=2).alignment = align_left

                for col in range(1, max_col + 1):
                    header = ws.cell(row=4, column=col).value
                    if header in percent_cols:
                        for row in range(5, max_row + 1): ws.cell(row=row, column=col).number_format = "0.00%"
                    else:
                        if col >= 3:
                            for row in range(5, max_row + 1): ws.cell(row=row, column=col).number_format = "#,##0"

                # Đổ màu cho dòng tổng cộng (TOTAL) ở dòng số 5
                for col in range(1, max_col + 1):
                    ws.cell(row=5, column=col).fill = fill_total; ws.cell(row=5, column=col).font = font_bold

                # Tự động gán độ rộng các cột cố định trực quan
                for col in range(1, max_col + 1):
                    if col == 1: ws.column_dimensions[get_column_letter(col)].width = 12
                    elif col == 2: ws.column_dimensions[get_column_letter(col)].width = 30
                    else: ws.column_dimensions[get_column_letter(col)].width = 18

                ws.freeze_panes = "C5"
                ws.auto_filter.ref = f"A4:{get_column_letter(max_col)}{max_row}"

                # Xuất file hoàn thiện ra ngoài cho người dùng
                final_output = io.BytesIO()
                wb.save(final_output)
                
                st.success("🎉 Hoàn tất định dạng báo cáo mỹ thuật theo quy chuẩn!")
                st.download_button("📥 Tải Báo Cáo Phân Tích Đã Tô Màu (.xlsx)", data=final_output.getvalue(), file_name=f"Bao_Cao_Thanh_Khoan_Thap_Màu_{thang_t}.xlsx")
        else:
            st.error("⚠️ Vui lòng cung cấp tệp dữ liệu CRM32 của cả 2 tháng!")
