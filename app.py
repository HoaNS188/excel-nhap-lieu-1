
import streamlit as st
import pandas as pd
import requests
from io import BytesIO

def get_onedrive_download_url(shared_url):
    if "1drv.ms" in shared_url:
        r = requests.get(shared_url, allow_redirects=True)
        return r.url.replace("redir?", "download?").replace("embed?", "download?")
    elif "onedrive.live.com" in shared_url:
        return shared_url.replace("redir?", "download?")
    else:
        return shared_url

st.set_page_config(page_title="Nhập liệu Excel OneDrive", layout="wide")
st.title("📋 Ứng dụng nhập liệu cho file Excel từ OneDrive")

default_url = "https://1drv.ms/x/c/ff3d3278bf1e16d9/EbcbAJKFsdtKifzIMel6nUgBb6SO8Z-LSWrYNiTSvVyP-Q?e=Lkp3et"
url = st.text_input("Nhập liên kết OneDrive chia sẻ (dạng 1drv.ms):", value=default_url)

if url:
    try:
        dl_url = get_onedrive_download_url(url)
        file = requests.get(dl_url).content

        # Đọc dữ liệu từ dòng tiêu đề thứ 5 (tức header=4)
        df = pd.read_excel(BytesIO(file), sheet_name="Bang CT", engine="openpyxl", header=4)

        # Lọc các cột có tên rõ ràng (loại bỏ cột Unnamed)
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

        st.success("✅ Đã tải dữ liệu từ sheet 'Bang CT'")
        st.dataframe(df.tail(10))

        st.markdown("### ➕ Nhập dữ liệu mới")
        input_data = {col: st.text_input(f"{col}:", "") for col in df.columns}

        if st.button("Thêm dòng"):
            df_new = df.append(input_data, ignore_index=True)
            excel_output = BytesIO()
            df_new.to_excel(excel_output, index=False, sheet_name="Bang CT")
            st.success("🎉 Dữ liệu đã được thêm!")
            st.download_button(
                label="📥 Tải file Excel mới",
                data=excel_output.getvalue(),
                file_name="Bang_CT_CapNhat.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ Đã xảy ra lỗi: {e}")
