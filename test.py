import streamlit as st
import streamlit_shadcn_ui as ui
import pandas as pd

# 使用者選擇檔案類型
file_type = ui.tabs(options=['型錄PDF', '貨號表', '商品資料表', '商品圖檔', '對照表'], default_value='型錄PDF', key="tabs")

# 根據選擇顯示對應的預覽內容與下載按鈕
if file_type == '型錄PDF':
    st.image("範例檔案/PDF範例.png")
    with open("範例檔案/型錄範例.pdf", "rb") as file:
        st.download_button(
            label="下載範例型錄PDF",
            data=file,
            file_name="catalog.pdf",
            mime="application/pdf"
        )

elif file_type == '貨號表':
    # 預覽 CSV 或 Excel 表格，並指定編碼格式
    st.write("### 範例貨號表")
    data = pd.read_csv("範例檔案/貨號表範例.csv", encoding='big5')  # 若是檔案使用 Big5 編碼，這裡需要明確指定
    st.dataframe(data)

    # 下載按鈕
    with open("範例檔案/貨號表範例.csv", "rb") as file:
        st.download_button(
            label="下載範例貨號表",
            data=file,
            file_name="item_list.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif file_type == '商品資料表':
    # 預覽商品資料表
    st.write("### 範例商品資料表")
    data = pd.read_csv("範例檔案/商品資料表範例.csv")
    st.dataframe(data)

    # 下載按鈕
    with open("範例檔案/商品資料表範例.csv", "rb") as file:
        st.download_button(
            label="下載範例商品資料表",
            data=file,
            file_name="product_data.csv",
            mime="text/csv"
        )

elif file_type == '商品圖檔':
    # 商品圖檔無法直接預覽，提示用戶
    st.write("### 商品圖檔範例")
    st.image("範例檔案/商品圖檔範例圖.png")
    st.warning("請以貨號為檔名，並打包成 Zip")
    
elif file_type == '對照表':
    # 預覽對照表
    st.write("### 範例對照表")
    data = pd.read_csv("範例檔案/對照表範例.csv", encoding='big5')
    st.dataframe(data)

    # 下載按鈕
    with open("範例檔案/對照表範例.csv", "rb") as file:
        st.download_button(
            label="下載範例對照表",
            data=file,
            file_name="comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
