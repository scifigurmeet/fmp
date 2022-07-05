import pandas as pd
from woocommerce import API
import concurrent.futures

def update_inventory(sku, stock):
    global wcapi
    global count
    data = wcapi.get("products?sku=" + sku).json()
    if len(data) != 0 and "id" in data[0]:
        data = data[0]
        productID = data['id']
        oldStock = data['stock_quantity']
        response = wcapi.put(f"products/{productID}", {
            "manage_stock": True,
            "stock_quantity": stock
        }).json()
        if "id" in response:
            newStock = response['stock_quantity']
            print(
                f'{sku} - Stock Quantity updated from {oldStock} to {newStock} - {round((count/total)*100, 2)}%'
            )
        else:
            print(
                f'{sku} - Stock Quantity update failed - {round((count/total)*100, 2)}%'
            )
            print(response)
    else:
        print(f'{sku} - Product not found - {round((count/total)*100, 2)}%')
    count += 1

wcapi = API(url="https://furnishmyplace.com",
            consumer_key="ck_b40fc3b7b657ce5f4c7bf64cc87755c0adc0f33e",
            consumer_secret="cs_bbecec3d5842fc6cfa3f1c1180cd99532be59cdc",
            version="wc/v3",
            timeout=100)

uploaded_file = "C:/Users/scifi/OneDrive/Desktop/butlerInventory.csv"

if uploaded_file is not None:
    count = 0
    df = pd.read_csv(uploaded_file)
    total = df.shape[0]
    if 'SKU' in df.columns and 'Stock' in df.columns:
        with concurrent.futures.ThreadPoolExecutor() as executor:
            executor.map(update_inventory, [row["SKU"] for index, row in df.iterrows()], [row["Stock"] for index, row in df.iterrows()])
    else:
        print("The data does not have the columns 'SKU' and 'Stock'")