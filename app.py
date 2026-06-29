import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =============================================================================
# CẤU HÌNH HỆ THỐNG & GIAO DIỆN
# =============================================================================
st.set_page_config(
    page_title="Hệ Thống Giám Sát Từ Xa Dữ Liệu Tín Dụng",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 Hệ Thống Xử Lý & Giám Sát Dữ Liệu Tín Dụng")
st.markdown("Ứng dụng tự động hóa tính toán dữ liệu tín dụng theo các tiêu chí kiểm soát rủi ro vận hành.")

# Hàm xử lý đọc file Excel thông minh chấp nhận cả .xls và .xlsx bản chuẩn hoặc bản giả lập HTML
def read_excel_smart(uploaded_files, target_month=None):
    if not uploaded_files:
        return pd.DataFrame()
    
    if not isinstance(uploaded_files, list):
        uploaded_files = [uploaded_files]
        
    list_df = []
    for f in uploaded_files:
        f_bytes = f.read()
        f.seek(0) # Reset con trỏ file
        df_temp = pd.DataFrame()
        
        try:
            # Thử cách 1: Dùng engine mặc định dựa trên đuôi file
            engine_pick = "xlrd" if f.name.endswith('.xls') else "openpyxl"
            df_temp = pd.read_excel(io.BytesIO(f_bytes), engine=engine_pick, dtype=str)
        except Exception:
            try:
                # Thử cách 2: Ép cấu trúc chéo engine đề phòng sai đuôi mở rộng
                engine_pick = "openpyxl" if f.name.endswith('.xls') else "xlrd"
                df_temp = pd.read_excel(io.BytesIO(f_bytes), engine=engine_pick, dtype=str)
            except Exception:
                try:
                    # Thử cách 3: Xử lý trường hợp file .xls giả lập từ bảng HTML
                    df_list = pd.read_html(io.BytesIO(f_bytes))
                    df_temp = df_list[0].astype(str) if df_list else pd.DataFrame()
                except Exception as e:
                    st.error(f"❌ Không thể đọc file {f.name}. Lỗi chi tiết: {str(e)}")
                    continue
                    
        if not df_temp.empty:
            df_temp["TEN_FILE"] = f.name
            if target_month:
                df_temp["THANG"] = target_month
            list_df.append(df_temp)
            
    if list_df:
        return pd.concat(list_df, ignore_index=True)
    return pd.DataFrame()

# =============================================================================
# PHÂN CHIA GIAO DIỆN CHỨC NĂNG THEO 4 TIÊU CHÍ
# =============================================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "1️⃣ Tiêu chí 1: Dư Nợ Cơ Cấu", 
    "2️⃣ Tiêu chí 2: Miễn Giảm Lãi", 
    "3️⃣ Tiêu chí 3: Lũy Kế Nợ Quá Hạn", 
    "4️⃣ Tiêu chí 4: Biến Động Thanh Khoản"
])

# =============================================================================
# TIÊU CHÍ 1: DƯ NỢ CƠ CẤU (GIỮ NGUYÊN LOGIC)
# =============================================================================
with tab1:
    st.header("1️⃣ Tính Toán Biến Động Dư Nợ Cơ Cấu")
    st.markdown("Đối chiếu chênh lệch danh sách khách hàng cơ cấu giữa tháng này và tháng trước kèm thông tin DPRR.")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📅 Dữ liệu Tháng này (T)")
        crm32_nay_files = st.file_uploader("Upload File(s) CRM32 Tháng này:", accept_multiple_files=True, key="tc1_32_t")
        crm4_nay_files = st.file_uploader("Upload File(s) CRM4 Tháng này:", accept_multiple_files=True, key="tc1_4_t")
    with col2:
        st.subheader("⏮️ Dữ liệu Tháng trước (T-1)")
        crm32_truoc_files = st.file_uploader("Upload File(s) CRM32 Tháng trước:", accept_multiple_files=True, key="tc1_32_t1")
        crm4_truoc_files = st.file_uploader("Upload File(s) CRM4 Tháng trước:", accept_multiple_files=True, key="tc1_4_t1")
        
    st.subheader("🛡️ File Dự Phòng Rủi Ro (DPRR)")
    sddp_file = st.file_uploader("Upload File DPRR_T5.xlsx:", type=["xlsx", "xls"], key="tc1_dprr")
    
    if st.button("🚀 Chạy Tiến Trình Tiêu Chí 1", key="btn_tc1"):
        if crm32_nay_files and crm4_nay_files and crm32_truoc_files and crm4_truoc_files and sddp_file:
            with st.spinner("Đang tính toán logic..."):
                df_crm32_nay = read_excel_smart(crm32_nay_files)
                df_crm4_nay = read_excel_smart(crm4_nay_files)
                df_crm32_truoc = read_excel_smart(crm32_truoc_files)
                df_crm4_truoc = read_excel_smart(crm4_truoc_files)
                df_sddp = read_excel_smart(sddp_file)
                
                # Logic xử lý tháng này
                df_crm32_cc_nay = df_crm32_nay[
                    df_crm32_nay["SCHM_DESC"].astype(str).str.upper().str.replace(r"\s+", " ", regex=True).str.contains("CO CAU", na=False)
                ].drop_duplicates(subset=["CUSTSEQLN"])
                
                df_join_nay = df_crm32_cc_nay.merge(df_crm4_nay, left_on="CUSTSEQLN", right_on="CIF_KH_VAY", how="inner")
                df_join_nay["DU_NO_PHAN_BO_QUY_DOI"] = pd.to_numeric(df_join_nay["DU_NO_PHAN_BO_QUY_DOI"], errors="coerce").fillna(0)
                
                df_kq_nay = df_join_nay.groupby(["CIF_KH_VAY", "SCHM_DESC", "NHOM_NO"], as_index=False)["DU_NO_PHAN_BO_QUY_DOI"].sum().rename(
                    columns={"DU_NO_PHAN_BO_QUY_DOI": "TONG_DU_NO_THANG_5", "NHOM_NO": "NHOM_NO_T5"}
                )
                df_kh = df_join_nay[["CIF_KH_VAY", "KHACH_HANG"]].drop_duplicates("CIF_KH_VAY")
                df_kq_nay = df_kq_nay.merge(df_kh, on="CIF_KH_VAY", how="left")
                
                # Logic xử lý tháng trước
                df_crm32_cc_truoc = df_crm32_truoc[
                    df_crm32_truoc["SCHM_DESC"].astype(str).str.upper().str.replace(r"\s+", " ", regex=True).str.contains("CO CAU", na=False)
                ].drop_duplicates(subset=["CUSTSEQLN"])
                
                df_join_truoc = df_crm32_cc_truoc.merge(df_crm4_truoc, left_on="CUSTSEQLN", right_on="CIF_KH_VAY", how="inner")
                df_join_truoc["DU_NO_PHAN_BO_QUY_DOI"] = pd.to_numeric(df_join_truoc["DU_NO_PHAN_BO_QUY_DOI"], errors="coerce").fillna(0)
                
                df_kq_truoc = df_join_truoc.groupby(["CIF_KH_VAY", "SCHM_DESC", "NHOM_NO"], as_index=False)["DU_NO_PHAN_BO_QUY_DOI"].sum().rename(
                    columns={"DU_NO_PHAN_BO_QUY_DOI": "TONG_DU_NO_THANG_4", "NHOM_NO": "NHOM_NO_T4"}
                )
                df_kh_trc = df_join_truoc[["CIF_KH_VAY", "KHACH_HANG"]].drop_duplicates("CIF_KH_VAY")
                df_kq_truoc = df_kq_truoc.merge(df_kh_trc, on="CIF_KH_VAY", how="left")
                
                # Đổi tên và Đối chiếu hai tháng
                df_kq_nay_rename = df_kq_nay.rename(columns={"CIF_KH_VAY": "CIF_KH_VAY_NAY", "SCHM_DESC": "SCHM_DESC_NAY", "KHACH_HANG": "TEN_KH_T5"})
                df_kq_truoc_rename = df_kq_truoc.rename(columns={"CIF_KH_VAY": "CIF_KH_VAY_TRUOC", "SCHM_DESC": "SCHM_DESC_TRUOC", "KHACH_HANG": "TEN_KH_T4"})
                
                df_doi_chieu = df_kq_nay_rename.merge(df_kq_truoc_rename, left_on=["CIF_KH_VAY_NAY", "SCHM_DESC_NAY"], right_on=["CIF_KH_VAY_TRUOC", "SCHM_DESC_TRUOC"], how="outer").fillna(0)
                df_doi_chieu["CHENH_LECH"] = df_doi_chieu["TONG_DU_NO_THANG_5"] - df_doi_chieu["TONG_DU_NO_THANG_4"]
                df_doi_chieu = df_doi_chieu.sort_values("CHENH_LECH", ascending=False)
                
                # Map thông tin DPRR
                df_sddp["PHAT_SINH_NO"] = pd.to_numeric(df_sddp["PHAT_SINH_NO"], errors="coerce").fillna(0)
                df_doi_chieu["CIF_KH_VAY_NAY"] = df_doi_chieu["CIF_KH_VAY_NAY"].astype(str).str.strip().str.replace(".0", "", regex=False)
                df_sddp["CIF"] = df_sddp["CIF"].astype(str).str.strip().str.replace(".0", "", regex=False)
                
                df_sddp_sum = df_sddp.groupby("CIF")["PHAT_SINH_NO"].sum()
                df_doi_chieu["DPRR"] = df_doi_chieu["CIF_KH_VAY_NAY"].map(df_sddp_sum)
                
                st.success("✅ Đã xử lý xong dữ liệu Tiêu chí 1!")
                st.dataframe(df_doi_chieu.head(100))
                
                out_buffer = io.BytesIO()
                df_doi_chieu.to_excel(out_buffer, index=False)
                st.download_button("📥 Tải File DU_NO_CO_CAU_T52026.xlsx", data=out_buffer.getvalue(), file_name="DU_NO_CO_CAU_T52026.xlsx")
        else:
            st.error("⚠️ Vui lòng tải lên đầy đủ các file theo yêu cầu để xử lý.")

# =============================================================================
# TIÊU CHÍ 2: MIỄN GIẢM LÃI (GIỮ NGUYÊN LOGIC)
# =============================================================================
with tab2:
    st.header("2️⃣ Báo Cáo Xử Lý Miễn Giảm Lãi Khách Hàng")
    st.markdown("Tổng hợp thông tin từ màn hình HLAWINT, doanh số thu nợ nội bảng, ngoại bảng và thu lãi.")
    
    f_hlawint = st.file_uploader("1. Chọn file Màn hình HLAWINT (Muc6_Manhinh_HLAWINT.xlsx):", type=["xlsx", "xls"])
    f_ngoaibang = st.file_uploader("2. Chọn file Thu ngoại bảng (ngoaibang.xlsx):", type=["xlsx", "xls"])
    f_noibang = st.file_uploader("3. Chọn file Thu gốc nội bảng (noibang.xlsx):", type=["xlsx", "xls"])
    f_thulai = st.file_uploader("4. Chọn file Thu lãi (thulai.xlsx):", type=["xlsx", "xls"])
    
    if st.button("🚀 Chạy Tiến Trình Tiêu Chí 2", key="btn_tc2"):
        if f_hlawint and f_ngoaibang and f_noibang and f_thulai:
            with st.spinner("Đang tổng hợp dữ liệu..."):
                df_goc = read_excel_smart(f_hlawint)
                cols_needed = ['SOL_ID', 'SOL_DESC', 'CIF_ID', 'CUST_NAME', 'INTAMT_VND', 'LY_DO']
                df_goc = df_goc[cols_needed]
                
                df_unique = df_goc.drop_duplicates(subset='CIF_ID')
                sum_intamt = df_goc.groupby('CIF_ID', as_index=False)['INTAMT_VND'].sum()
                
                result = pd.merge(sum_intamt, df_unique[['CIF_ID', 'SOL_ID', 'SOL_DESC', 'CUST_NAME', 'LY_DO']], on='CIF_ID', how='left')
                
                dprr = read_excel_smart(f_ngoaibang)
                dprr_sum = dprr.groupby('CIF 9 SO', as_index=False)['Số tiền thu'].sum().rename(columns={'Số tiền thu': 'Thu_ngoai_bang', 'CIF 9 SO': 'CIF_ID'})
                result = result.merge(dprr_sum, on='CIF_ID', how='left')
                
                thu_goc = read_excel_smart(f_noibang)
                thu_goc_sum = thu_goc.groupby('CIF 9 SO', as_index=False)['Số tiền thu'].sum().rename(columns={'Số tiền thu': 'Thu_goc_trong_thang', 'CIF 9 SO': 'CIF_ID'})
                result = result.merge(thu_goc_sum, on='CIF_ID', how='left')
                
                thu_lai = read_excel_smart(f_thulai)
                thu_lai_sum = thu_lai.groupby('CIF 9 SO', as_index=False)['Số tiền thu'].sum().rename(columns={'Số tiền thu': 'Thu_lai_trong_thang', 'CIF 9 SO': 'CIF_ID'})
                result = result.merge(thu_lai_sum, on='CIF_ID', how='left')
                
                result['Tong_thu'] = result['Thu_goc_trong_thang'].fillna(0) + result['Thu_lai_trong_thang'].fillna(0) + result['Thu_ngoai_bang'].fillna(0)
                
                new_cols = ['Thu_thieu_lai_trong_han', 'Mien_goc', 'Mien_lai_trong_han', 'Mien_lai_qua_han', 'Mien_lai_phat', 'Ghi_chu', 'Con_ton', 'Ty_quan', 'So_thu_phi', 'note']
                for c in new_cols:
                    result[c] = 0
                    
                result['Tham_quyen_phe_duyet'] = result['LY_DO']
                
                muc6 = df_goc.groupby('CIF_ID')['INTAMT_VND'].sum().reset_index()
                muc6['muc6'] = muc6['INTAMT_VND']
                result = result.merge(muc6[['CIF_ID', 'muc6']], on='CIF_ID', how='left')
                
                result['check_tongMGL'] = (result['Mien_goc'] + result['Mien_lai_trong_han'] + result['Mien_lai_qua_han'] + result['Mien_lai_phat']) - result['muc6']
                
                final_cols = [
                    'SOL_ID', 'SOL_DESC', 'CIF_ID', 'CUST_NAME', 'Tong_thu', 'Thu_ngoai_bang', 'Thu_goc_trong_thang',
                    'Thu_lai_trong_thang', 'Thu_thieu_lai_trong_han', 'Mien_goc', 'Mien_lai_trong_han', 'Mien_lai_qua_han',
                    'Mien_lai_phat', 'Tham_quyen_phe_duyet', 'Ghi_chu', 'Con_ton', 'Ty_quan', 'So_thu_phi', 'muc6', 'check_tongMGL', 'note'
                ]
                result = result[final_cols]
                
                st.success("✅ Đã xử lý xong dữ liệu Tiêu chí 2!")
                st.dataframe(result.head(100))
                
                out_buffer = io.BytesIO()
                result.to_excel(out_buffer, index=False)
                st.download_button("📥 Tải File XulyNo_T5.xlsx", data=out_buffer.getvalue(), file_name="XulyNo_T5.xlsx")
        else:
            st.error("⚠️ Vui lòng tải đầy đủ 4 file dữ liệu gốc.")

# =============================================================================
# TIÊU CHÍ 3: LŨY KẾ NỢ QUÁ HẠN (SỬ DỤNG FILE UPLOADER ĐỘNG)
# =============================================================================
with tab3:
    st.header("3️⃣ Tính Toán Lũy Kế Giải Ngân Quá Hạn")
    st.markdown("Cấu hình dải tháng và đăng tải trực tiếp các tệp tin hệ thống CRM32 tương ứng lên ứng dụng web.")
    
    col_t3_1, col_t3_2 = st.columns(2)
    with col_t3_1:
        start_m = st.text_input("Tháng bắt đầu (YYYY-MM):", "2026-01", key="tc3_start")
    with col_t3_2:
        end_m = st.text_input("Tháng kết thúc (YYYY-MM):", "2026-05", key="tc3_end")
        
    ds_thang = pd.period_range(start=start_m, end=end_m, freq="M").astype(str).tolist()
    
    st.markdown("---")
    st.write("📂 **Vui lòng chọn nạp tệp tin CRM32 cho từng tháng bên dưới:**")
    
    dict_uploaded_t3 = {}
    # Tạo các slot upload file động trên web dựa theo cấu hình chuỗi tháng ở trên
    for m in ds_thang:
        dict_uploaded_t3[m] = st.file_uploader(f"Chọn file(s) CRM32 cho Tháng {m}:", accept_multiple_files=True, key=f"t3_upload_{m}")
        
    if st.button("🚀 Chạy Tiến Trình Tiêu Chí 3", key="btn_tc3"):
        all_month_data = []
        is_missing_file = False
        
        for m in ds_thang:
            if not dict_uploaded_t3[m]:
                st.warning(f"⚠️ Thiếu tệp tin CRM32 đầu vào của tháng {m}")
                is_missing_file = True
            else:
                df_m_temp = read_excel_smart(dict_uploaded_t3[m], target_month=m)
                if not df_m_temp.empty:
                    all_month_data.append(df_m_temp)
                    
        if not is_missing_file and all_month_data:
            with st.spinner("Đang tính toán chuỗi dữ liệu lũy kế..."):
                df = pd.concat(all_month_data, ignore_index=True)
                df.columns = df.columns.str.strip().str.upper()
                
                col_brcd, col_chi_nhanh, col_du_no, col_nhom_no, col_ngay_giai_ngan = "BRCD", "CHI_NHANH", "DU_NO_QUY_DOI", "NHOM_NO_THEO_CIF", "NGAY_GIAI_NGAN"
                
                df[col_brcd] = df[col_brcd].astype(str).str.strip()
                df[col_chi_nhanh] = df[col_chi_nhanh].astype(str).str.strip()
                df[col_du_no] = pd.to_numeric(df[col_du_no], errors="coerce").fillna(0)
                df[col_nhom_no] = pd.to_numeric(df[col_nhom_no], errors="coerce")
                df[col_ngay_giai_ngan] = pd.to_datetime(df[col_ngay_giai_ngan], errors="coerce", dayfirst=True)
                df["THANG"] = df["THANG"].astype(str)
                
                def tinh_giai_ngan_qua_han_1_thang(data, t_str):
                    df_thang = data[data["THANG"] == t_str].copy()
                    ngay_bao_cao = pd.to_datetime(t_str + "-01") + pd.offsets.MonthEnd(0)
                    df_thang["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] = (ngay_bao_cao - df_thang[col_ngay_giai_ngan]).dt.days
                    df_qh = df_thang[df_thang[col_nhom_no].isin([2, 3, 4, 5])].copy()
                    
                    g1 = df_qh[(df_qh["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] >= 0) & (df_qh["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] <= 90)].groupby([col_brcd, col_chi_nhanh], as_index=False).agg(GIAI_NGAN_3_THANG_QH=(col_du_no, "sum"))
                    g2 = df_qh[(df_qh["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] >= 0) & (df_qh["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] <= 180)].groupby([col_brcd, col_chi_nhanh], as_index=False).agg(GIAI_NGAN_6_THANG_QH=(col_du_no, "sum"))
                    base = df_thang.groupby([col_brcd, col_chi_nhanh], as_index=False).agg(TONG_DU_NO_THANG=(col_du_no, "sum"))
                    
                    res = base.merge(g1, on=[col_brcd, col_chi_nhanh], how="left").merge(g2, on=[col_brcd, col_chi_nhanh], how="left").fillna(0)
                    res["THANG"] = t_str
                    res["NGAY_BAO_CAO"] = ngay_bao_cao
                    return res
                    
                ds_bc = [tinh_giai_ngan_qua_han_1_thang(df, t) for t in ds_thang]
                bao_cao_all_month = pd.concat(ds_bc, ignore_index=True)
                bao_cao_all_month["THANG_DATE"] = pd.to_datetime(bao_cao_all_month["THANG"] + "-01")
                bao_cao_all_month = bao_cao_all_month.sort_values(by=[col_brcd, "THANG_DATE"])
                
                bao_cao_all_month["LUY_KE_GIAI_NGAN_3_THANG_QH"] = bao_cao_all_month.groupby([col_brcd, col_chi_nhanh"])["GIAI_NGAN_3_THANG_QH"].cumsum()
                bao_cao_all_month["LUY_KE_GIAI_NGAN_6_THANG_QH"] = bao_cao_all_month.groupby([col_brcd, col_chi_nhanh"])["GIAI_NGAN_6_THANG_QH"].cumsum()
                
                pivot_luy_ke = bao_cao_all_month.pivot_table(index=[col_brcd, col_chi_nhanh], columns="THANG", values=["GIAI_NGAN_3_THANG_QH", "GIAI_NGAN_6_THANG_QH", "LUY_KE_GIAI_NGAN_3_THANG_QH", "LUY_KE_GIAI_NGAN_6_THANG_QH"], aggfunc="sum")
                pivot_luy_ke.columns = [f"{chi_tieu}_{t}" for chi_tieu, t in pivot_luy_ke.columns]
                pivot_luy_ke = pivot_luy_ke.reset_index()
                
                st.success("✅ Đã xử lý xong báo cáo lũy kế chéo nhiều tháng từ file upload!")
                st.dataframe(pivot_luy_ke.head(50))
                
                out_buffer = io.BytesIO()
                with pd.ExcelWriter(out_buffer, engine="openpyxl") as writer:
                    bao_cao_all_month.to_excel(writer, sheet_name="Data luy ke theo thang", index=False)
                    pivot_luy_ke.to_excel(writer, sheet_name="Bang ngang luy ke", index=False)
                    
                st.download_button("📥 Tải Xuống File Kết Quả Lũy Kế (.xlsx)", data=out_buffer.getvalue(), file_name=f"luy_ke_nhieu_thang_CRM32_{start_m}_den_{end_m}.xlsx")

# =============================================================================
# TIÊU CHÍ 4: THANH KHOẢN THẤP & CHUẨN MỸ THUẬT (SỬ DỤNG FILE UPLOADER TRỰC TIẾP)
# =============================================================================
with tab4:
    st.header("4️⃣ So Sánh Biến Động Thanh Khoản Thấp & Định Dạng Mỹ Thuật")
    st.markdown("Đăng tải trực tiếp file báo cáo tháng T và T-1 để đối chiếu dữ liệu phân tích biến động thanh khoản.")
    
    thang_t_input = st.text_input("Nhập cấu hình chuỗi tháng báo cáo T (YYYY-MM):", "2026-04", key="tc4_t")
    
    thang_t_date = pd.to_datetime(thang_t_input + "-01")
    thang_t_1_date = thang_t_date - pd.DateOffset(months=1)
    thang_t_str = thang_t_date.strftime("%Y-%m")
    thang_t_1_str = thang_t_1_date.strftime("%Y-%m")
    
    col_t4_1, col_t4_2 = st.columns(2)
    with col_t4_1:
        files_tc4_t = st.file_uploader(f"Chọn file(s) CRM32 cho Tháng T ({thang_t_str}):", accept_multiple_files=True, key="tc4_upload_t")
    with col_t4_2:
        files_tc4_t1 = st.file_uploader(f"Chọn file(s) CRM32 cho Tháng T-1 ({thang_t_1_str}):", accept_multiple_files=True, key="tc4_upload_t1")
        
    if st.button("🚀 Chạy Tiến Trình Tiêu Chí 4 & Xuất Bản", key="btn_tc4"):
        if files_tc4_t and files_tc4_t1:
            with st.spinner("Hệ thống đang chạy logic đối chiếu và vẽ biểu mẫu Excel..."):
                df_t_1 = read_excel_smart(files_tc4_t1, target_month=thang_t_1_str)
                df_t = read_excel_smart(files_tc4_t, target_month=thang_t_str)
                
                df = pd.concat([df_t_1, df_t], ignore_index=True)
                df.columns = df.columns.str.strip().str.upper()
                
                col_brcd, col_chi_nhanh, col_du_no, col_nhom_no, col_ngay_giai_ngan = "BRCD", "CHI_NHANH", "DU_NO_QUY_DOI", "NHOM_NO_THEO_CIF", "NGAY_GIAI_NGAN"
                df[col_brcd] = df[col_brcd].astype(str).str.strip()
                df[col_chi_nhanh] = df[col_chi_nhanh].astype(str).str.strip()
                df[col_du_no] = pd.to_numeric(df[col_du_no], errors="coerce").fillna(0)
                df[col_nhom_no] = pd.to_numeric(df[col_nhom_no], errors="coerce")
                df[col_ngay_giai_ngan] = pd.to_datetime(df[col_ngay_giai_ngan], errors="coerce")
                df["THANG"] = df["THANG"].astype(str)
                
                def tinh_chi_tieu_1_thang(data, t_val):
                    df_thang = data[data["THANG"] == t_val].copy()
                    ngay_bao_cao = pd.to_datetime(t_val + "-01") + pd.offsets.MonthEnd(0)
                    df_thang["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] = (ngay_bao_cao - df_thang[col_ngay_giai_ngan]).dt.days
                    
                    du_no = df_thang.groupby([col_brcd, col_chi_nhanh], as_index=False).agg(DU_NO=(col_du_no, "sum"))
                    df_qh = df_thang[df_thang[col_nhom_no].isin([2, 3, 4, 5])].copy()
                    du_no_qh = df_qh.groupby([col_brcd, col_chi_nhanh], as_index=False).agg(DU_NO_QH=(col_du_no, "sum"))
                    
                    df_xau = df_thang[df_thang[col_nhom_no].isin([3, 4, 5])].copy()
                    du_no_xau = df_xau.groupby([col_brcd, col_chi_nhanh], as_index=False).agg(DU_NO_XAU=(col_du_no, "sum"))
                    
                    g3 = df_thang[(df_thang["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] >= 0) & (df_thang["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] <= 90) & (df_thang[col_nhom_no].isin([2, 3, 4, 5]))].groupby([col_brcd, col_chi_nhanh], as_index=False).agg(GIAI_NGAN_3_THANG_QH=(col_du_no, "sum"))
                    g6 = df_thang[(df_thang["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] >= 0) & (df_thang["SO_NGAY_TU_GIAI_NGAN_DEN_BAO_CAO"] <= 180) & (df_thang[col_nhom_no].isin([2, 3, 4, 5]))].groupby([col_brcd, col_chi_nhanh], as_index=False).agg(GIAI_NGAN_6_THANG_QH=(col_du_no, "sum"))
                    
                    res = du_no.merge(du_no_qh, on=[col_brcd, col_chi_nhanh], how="left").merge(du_no_xau, on=[col_brcd, col_chi_nhanh], how="left").merge(g3, on=[col_brcd, col_chi_nhanh], how="left").merge(g6, on=[col_brcd, col_chi_nhanh], how="left").fillna(0)
                    res["TY_LE_QH"] = np.where(res["DU_NO"] != 0, res["DU_NO_QH"] / res["DU_NO"], 0)
                    res["TY_LE_NPL"] = np.where(res["DU_NO"] != 0, res["DU_NO_XAU"] / res["DU_NO"], 0)
                    res["THANG"] = t_val
                    res["NGAY_BAO_CAO"] = ngay_bao_cao
                    return res
                    
                bc_t_1 = tinh_chi_tieu_1_thang(df, thang_t_1_str)
                bc_t = tinh_chi_tieu_1_thang(df, thang_t_str)
                
                bc_t_1_rename = bc_t_1.rename(columns={"DU_NO": "DU_NO_T_1", "DU_NO_QH": "DU_NO_QH_T_1", "DU_NO_XAU": "DU_NO_XAU_T_1", "TY_LE_QH": "TY_LE_QH_T_1", "TY_LE_NPL": "TY_LE_NPL_T_1", "GIAI_NGAN_3_THANG_QH": "GIAI_NGAN_3_THANG_QH_T_1", "GIAI_NGAN_6_THANG_QH": "GIAI_NGAN_6_THANG_QH_T_1"})
                bc_t_rename = bc_t.rename(columns={"DU_NO": "DU_NO_T", "DU_NO_QH": "DU_NO_QH_T", "DU_NO_XAU": "DU_NO_XAU_T", "TY_LE_QH": "TY_LE_QH_T", "TY_LE_NPL": "TY_LE_NPL_T", "GIAI_NGAN_3_THANG_QH": "GIAI_NGAN_3_THANG_QH_T", "GIAI_NGAN_6_THANG_QH": "GIAI_NGAN_6_THANG_QH_T"})
                
                compare = bc_t_1_rename[[col_brcd, col_chi_nhanh, "DU_NO_T_1", "DU_NO_QH_T_1", "DU_NO_XAU_T_1", "TY_LE_QH_T_1", "TY_LE_NPL_T_1", "GIAI_NGAN_3_THANG_QH_T_1", "GIAI_NGAN_6_THANG_QH_T_1"]].merge(
                    bc_t_rename[[col_brcd, col_chi_nhanh, "DU_NO_T", "DU_NO_QH_T", "DU_NO_XAU_T", "TY_LE_QH_T", "TY_LE_NPL_T", "GIAI_NGAN_3_THANG_QH_T", "GIAI_NGAN_6_THANG_QH_T"]], on=[col_brcd, col_chi_nhanh], how="outer"
                ).fillna(0)
                
                compare["SS_DU_NO"] = compare["DU_NO_T"] - compare["DU_NO_T_1"]
                compare["SS_QH"] = compare["DU_NO_QH_T"] - compare["DU_NO_QH_T_1"]
                compare["SS_NPL"] = compare["DU_NO_XAU_T"] - compare["DU_NO_XAU_T_1"]
                compare["SS_TY_LE_QH"] = compare["TY_LE_QH_T"] - compare["TY_LE_QH_T_1"]
                compare["SS_TY_LE_NPL"] = compare["TY_LE_NPL_T"] - compare["TY_LE_NPL_T_1"]
                compare["SS_GIAI_NGAN_3_THANG_QH"] = compare["GIAI_NGAN_3_THANG_QH_T"] - compare["GIAI_NGAN_3_THANG_QH_T_1"]
                compare["SS_GIAI_NGAN_6_THANG_QH"] = compare["GIAI_NGAN_6_THANG_QH_T"] - compare["GIAI_NGAN_6_THANG_QH_T_1"]
                
                final_df = compare[[
                    col_brcd, col_chi_nhanh, "DU_NO_T_1", "DU_NO_QH_T_1", "DU_NO_XAU_T_1", "TY_LE_QH_T_1", "TY_LE_NPL_T_1", "GIAI_NGAN_3_THANG_QH_T_1", "GIAI_NGAN_6_THANG_QH_T_1",
                    "DU_NO_T", "DU_NO_QH_T", "DU_NO_XAU_T", "TY_LE_QH_T", "TY_LE_NPL_T", "GIAI_NGAN_3_THANG_QH_T", "GIAI_NGAN_6_THANG_QH_T",
                    "SS_DU_NO", "SS_QH", "SS_NPL", "SS_TY_LE_QH", "SS_TY_LE_NPL", "SS_GIAI_NGAN_3_THANG_QH", "SS_GIAI_NGAN_6_THANG_QH"
                ]].copy()
                
                final_df = final_df.rename(columns={
                    col_brcd: "BRCD", col_chi_nhanh: "CHI_NHANH",
                    "DU_NO_T_1": "Dư nợ_T-1", "DU_NO_QH_T_1": "Dư nợ QH_T-1", "DU_NO_XAU_T_1": "Dư nợ xấu_T-1", "TY_LE_QH_T_1": "%QH_T-1", "TY_LE_NPL_T_1": "%NPL_T-1", "GIAI_NGAN_3_THANG_QH_T_1": "Giải ngân 3 tháng quá hạn_T-1", "GIAI_NGAN_6_THANG_QH_T_1": "Giải ngân 6 tháng quá hạn_T-1",
                    "DU_NO_T": "Dư nợ_T", "DU_NO_QH_T": "Dư nợ QH_T", "DU_NO_XAU_T": "Dư nợ xấu_T", "TY_LE_QH_T": "%QH_T", "TY_LE_NPL_T": "%NPL_T", "GIAI_NGAN_3_THANG_QH_T": "Giải ngân 3 tháng quá hạn_T", "GIAI_NGAN_6_THANG_QH_T": "Giải ngân 6 tháng quá hạn_T",
                    "SS_DU_NO": "Dư nợ", "SS_QH": "QH", "SS_NPL": "NPL", "SS_TY_LE_QH": "%QH", "SS_TY_LE_NPL": "%NPL", "SS_GIAI_NGAN_3_THANG_QH": "Giải ngân 3 tháng quá hạn", "SS_GIAI_NGAN_6_THANG_QH": "Giải ngân 6 tháng quá hạn"
                })
                
                total_row = {"BRCD": "TOTAL", "CHI_NHANH": "Tổng cộng"}
                money_cols = ["Dư nợ_T-1", "Dư nợ QH_T-1", "Dư nợ xấu_T-1", "Giải ngân 3 tháng quá hạn_T-1", "Giải ngân 6 tháng quá hạn_T-1", "Dư nợ_T", "Dư nợ QH_T", "Dư nợ xấu_T", "Giải ngân 3 tháng quá hạn_T", "Giải ngân 6 tháng quá hạn_T", "Dư nợ", "QH", "NPL", "Giải ngân 3 tháng quá hạn", "Giải ngân 6 tháng quá hạn"]
                for m_c in money_cols:
                    total_row[m_c] = final_df[m_c].sum()
                    
                total_row["%QH_T-1"] = total_row["Dư nợ QH_T-1"] / total_row["Dư nợ_T-1"] if total_row["Dư nợ_T-1"] != 0 else 0
                total_row["%NPL_T-1"] = total_row["Dư nợ xấu_T-1"] / total_row["Dư nợ_T-1"] if total_row["Dư nợ_T-1"] != 0 else 0
                total_row["%QH_T"] = total_row["Dư nợ QH_T"] / total_row["Dư nợ_T"] if total_row["Dư nợ_T"] != 0 else 0
                total_row["%NPL_T"] = total_row["Dư nợ xấu_T"] / total_row["Dư nợ_T"] if total_row["Dư nợ_T"] != 0 else 0
                total_row["%QH"] = total_row["%QH_T"] - total_row["%QH_T-1"]
                total_row["%NPL"] = total_row["%NPL_T"] - total_row["%NPL_T-1"]
                
                final_df = pd.concat([pd.DataFrame([total_row]), final_df], ignore_index=True)
                
                wb = Workbook()
                ws = wb.active
                ws.title = "Tong hop"
                ws.views.sheetView[0].showGridLines = True
                
                from openpyxl.utils.dataframe import dataframe_to_rows
                for row_data in dataframe_to_rows(final_df, index=False, header=True):
                    ws.append(row_data)
                    
                max_row, max_col = ws.max_row, ws.max_column
                
                fill_yellow = PatternFill("solid", fgColor="FFFF00")
                fill_green = PatternFill("solid", fgColor="C6E0B4")
                fill_gray = PatternFill("solid", fgColor="D9D9D9")
                fill_total = PatternFill("solid", fgColor="A9D08E")
                font_red = Font(name="Arial", color="FF0000", bold=True)
                font_bold = Font(name="Arial", bold=True)
                align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
                align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                
                ws.insert_rows(1, 3)
                
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
                ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=9)
                ws.cell(row=1, column=3).value = f"Tháng T-1 ({thang_t_1_str})"
                ws.merge_cells(start_row=1, start_column=10, end_row=1, end_column=16)
                ws.cell(row=1, column=10).value = f"Tháng T ({thang_t_str})"
                ws.merge_cells(start_row=1, start_column=17, end_row=1, end_column=23)
                ws.cell(row=1, column=17).value = "So sánh (T) - (T-1)"
                
                for col in range(3, 24):
                    ws.cell(row=2, column=col).value = col - 2
                    
                for row in [1, 2, 4]:
                    for col in range(1, max_col + 1):
                        cell = ws.cell(row=row, column=col)
                        cell.alignment = align_center
                        cell.border = thin_border
                        cell.font = font_bold
                        
                for col in range(3, 17):
                    ws.cell(row=1, column=col).fill = fill_yellow
                    ws.cell(row=2, column=col).fill = fill_yellow
                    ws.cell(row=4, column=col).fill = fill_yellow
                for col in range(17, 24):
                    ws.cell(row=1, column=col).fill = fill_green
                    ws.cell(row=2, column=col).fill = fill_green
                    ws.cell(row=4, column=col).fill = fill_green
                for col in range(1, 3):
                    ws.cell(row=4, column=col).fill = fill_gray
                    
                ws.cell(row=1, column=3).font = font_red
                ws.cell(row=1, column=10).font = font_red
                ws.cell(row=1, column=17).font = font_red
                
                for row in range(5, ws.max_row + 1):
                    for col in range(1, max_col + 1):
                        cell = ws.cell(row=row, column=col)
                        cell.border = thin_border
                        cell.alignment = align_center
                        
                for row in range(5, ws.max_row + 1):
                    ws.cell(row=row, column=2).alignment = align_left
                    
                percent_cols = ["%QH_T-1", "%NPL_T-1", "%QH_T", "%NPL_T", "%QH", "%NPL"]
                for col in range(1, max_col + 1):
                    header_val = ws.cell(row=4, column=col).value
                    if header_val in percent_cols:
                        for row in range(5, ws.max_row + 1):
                            ws.cell(row=row, column=col).number_format = "0.00%"
                    else:
                        if col >= 3:
                            for row in range(5, ws.max_row + 1):
                                ws.cell(row=row, column=col).number_format = "#,##0"
                                
                for col in range(1, max_col + 1):
                    ws.cell(row=5, column=col).fill = fill_total
                    ws.cell(row=5, column=col).font = font_bold
                    
                for col in range(1, max_col + 1):
                    if col == 1: ws.column_dimensions[get_column_letter(col)].width = 12
                    elif col == 2: ws.column_dimensions[get_column_letter(col)].width = 32
                    else: ws.column_dimensions[get_column_letter(col)].width = 18
                    
                ws.freeze_panes = "C6"
                ws.auto_filter.ref = f"A4:{get_column_letter(max_col)}{ws.max_row}"
                
                out_art = io.BytesIO()
                wb.save(out_art)
                
                st.success("🎉 Đã đối chiếu xong dữ liệu và định dạng biểu mẫu thành công!")
                st.dataframe(final_df.head(20))
                
                st.download_button(
                    label="📥 Tải File Báo Cáo Định Dạng Mỹ Thuật (.XLSX)",
                    data=out_art.getvalue(),
                    file_name=f"so_sanh_CRM32_{thang_t_str}_vs_{thang_t_1_str}_my_thuat.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("⚠️ Vui lòng tải đầy đủ file CRM32 của cả 2 tháng để hệ thống đối chiếu.")
