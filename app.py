import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Đối chiếu dư nợ cơ cấu",
    layout="wide"
)

st.title("Đối chiếu dư nợ cơ cấu T4 - T5")

st.markdown("---")

col1,col2=st.columns(2)

with col1:

    st.subheader("Tháng hiện tại")

    crm32_nay=st.file_uploader(
        "CRM32",
        type=["xls","xlsx"],
        accept_multiple_files=True,
        key="crm32nay"
    )

    crm4_nay=st.file_uploader(
        "CRM4",
        type=["xls","xlsx"],
        accept_multiple_files=True,
        key="crm4nay"
    )

with col2:

    st.subheader("Tháng trước")

    crm32_truoc=st.file_uploader(
        "CRM32 ",
        type=["xls","xlsx"],
        accept_multiple_files=True,
        key="crm32truoc"
    )

    crm4_truoc=st.file_uploader(
        "CRM4 ",
        type=["xls","xlsx"],
        accept_multiple_files=True,
        key="crm4truoc"
    )

st.markdown("---")

dprr_file=st.file_uploader(
    "Upload DPRR",
    type=["xlsx"]
)

if st.button("Thực hiện"):

    if (
        len(crm32_nay)==0
        or len(crm4_nay)==0
        or len(crm32_truoc)==0
        or len(crm4_truoc)==0
        or dprr_file is None
    ):
        st.error("Thiếu file")
        st.stop()

    def read_excel(files):

        dfs=[]

        for f in files:

            if f.name.endswith(".xls"):
                df=pd.read_excel(
                    f,
                    engine="xlrd",
                    dtype=str
                )
            else:
                df=pd.read_excel(
                    f,
                    engine="openpyxl",
                    dtype=str
                )

            dfs.append(df)

        return pd.concat(dfs,ignore_index=True)


    df_crm32_nay=read_excel(crm32_nay)
    df_crm4_nay=read_excel(crm4_nay)

    df_crm32_truoc=read_excel(crm32_truoc)
    df_crm4_truoc=read_excel(crm4_truoc)

    #######################################################
    ## Hàm xử lý
    #######################################################

    def xu_ly(df32,df4,thang):

        df32=df32.copy()
        df4=df4.copy()

        df_cc=(
            df32[
                df32["SCHM_DESC"]
                .astype(str)
                .str.upper()
                .str.replace(r"\s+"," ",regex=True)
                .str.contains("CO CAU",na=False)
            ]
            .drop_duplicates(subset=["CUSTSEQLN"])
        )

        df_join=df_cc.merge(
            df4,
            left_on="CUSTSEQLN",
            right_on="CIF_KH_VAY",
            how="inner"
        )

        df_join["DU_NO_PHAN_BO_QUY_DOI"]=pd.to_numeric(
            df_join["DU_NO_PHAN_BO_QUY_DOI"],
            errors="coerce"
        ).fillna(0)

        df=(
            df_join
            .groupby(
                ["CIF_KH_VAY","SCHM_DESC","NHOM_NO"],
                as_index=False
            )["DU_NO_PHAN_BO_QUY_DOI"]
            .sum()
        )

        df=df.rename(
            columns={
                "DU_NO_PHAN_BO_QUY_DOI":f"TONG_DU_NO_THANG_{thang}",
                "NHOM_NO":f"NHOM_NO_T{thang}"
            }
        )

        kh=df_join[
            ["CIF_KH_VAY","KHACH_HANG"]
        ].drop_duplicates("CIF_KH_VAY")

        df=df.merge(
            kh,
            on="CIF_KH_VAY",
            how="left"
        )

        return df

    df_t5=xu_ly(df_crm32_nay,df_crm4_nay,5)
    df_t4=xu_ly(df_crm32_truoc,df_crm4_truoc,4)

    ###################################################

    df_t5=df_t5.rename(columns={
        "CIF_KH_VAY":"CIF_KH_VAY_NAY",
        "SCHM_DESC":"SCHM_DESC_NAY",
        "KHACH_HANG":"TEN_KH_T5"
    })

    df_t4=df_t4.rename(columns={
        "CIF_KH_VAY":"CIF_KH_VAY_TRUOC",
        "SCHM_DESC":"SCHM_DESC_TRUOC",
        "KHACH_HANG":"TEN_KH_T4"
    })

    df=df_t5.merge(
        df_t4,
        left_on=[
            "CIF_KH_VAY_NAY",
            "SCHM_DESC_NAY"
        ],
        right_on=[
            "CIF_KH_VAY_TRUOC",
            "SCHM_DESC_TRUOC"
        ],
        how="outer"
    )

    df=df.fillna(0)

    df["CHENH_LECH"]=(
        df["TONG_DU_NO_THANG_5"]
        -df["TONG_DU_NO_THANG_4"]
    )

    ###################################################
    ## DPRR
    ###################################################

    dprr=pd.read_excel(
        dprr_file,
        dtype={"CIF":str}
    )

    dprr["PHAT_SINH_NO"]=pd.to_numeric(
        dprr["PHAT_SINH_NO"],
        errors="coerce"
    ).fillna(0)

    df["CIF_KH_VAY_NAY"]=(
        df["CIF_KH_VAY_NAY"]
        .astype(str)
        .str.replace(".0","",regex=False)
        .str.strip()
    )

    dprr["CIF"]=(
        dprr["CIF"]
        .astype(str)
        .str.replace(".0","",regex=False)
        .str.strip()
    )

    map_dprr=dprr.groupby("CIF")["PHAT_SINH_NO"].sum()

    df["DPRR"]=df["CIF_KH_VAY_NAY"].map(map_dprr)

    df=df.sort_values(
        "CHENH_LECH",
        ascending=False
    )

    st.success("Hoàn thành")

    st.dataframe(df.head(100),use_container_width=True)

    excel=df.to_excel(index=False)

    import io

    output=io.BytesIO()

    with pd.ExcelWriter(output,engine="openpyxl") as writer:
        df.to_excel(writer,index=False)

    st.download_button(
        "Tải Excel",
        data=output.getvalue(),
        file_name="DU_NO_CO_CAU.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
