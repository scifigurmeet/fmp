from itertools import product
import streamlit as st
import pandas as pd
from woocommerce import API
import concurrent.futures

def update_inventory(sku, stock):
    data = wcapi.get("products?sku=" + sku).json()[0]
    productID = data['id']
    oldStock = data['stock_quantity']
    response = wcapi.put(f"products/{productID}", {
        "manage_stock": True,
        "stock_quantity": stock
    }).json()
    if "id" in response:
        newStock = response['stock_quantity']
        print(f'{sku} - Stock Quantity updated from {oldStock} to {newStock}')
        st.write(
            f'{sku} - Stock Quantity updated from {oldStock} to {newStock}')
    else:
        print(f'{sku} - Stock Quantity update failed')
        print(response)


st.header("FurnishMyPlace Inventory Updates")

password = st.text_input("Enter Password To Continue", type="password")

if password != "fmp%*8m68Fn":
    st.error("Please Enter the Correct Password To Continue...")
else:
    wcapi = API(url="https://furnishmyplace.com",
                consumer_key="ck_b40fc3b7b657ce5f4c7bf64cc87755c0adc0f33e",
                consumer_secret="cs_bbecec3d5842fc6cfa3f1c1180cd99532be59cdc",
                version="wc/v3",
                timeout=100)

    uploaded_file = st.file_uploader("Upload Inventory CSV File", type="csv")

    if uploaded_file is not None:
        if st.button("Process Inventory"):
            df = pd.read_csv(uploaded_file)
            if 'SKU' in df.columns and 'Stock' in df.columns:
                with concurrent.futures.ThreadPoolExecutor() as executor:
                    for sku, stock in product(df['SKU'], df['Stock']):
                        executor.submit(update_inventory, sku, stock)
            else:
                st.error("The data does not have the columns 'SKU' and 'Stock'")