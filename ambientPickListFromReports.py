import streamlit as st
import pandas as pd
from io import BytesIO
import base64
import time
import math


st.set_page_config(layout="wide")
st.write("""# Ambient Pick List from Reports""")

number = st.number_input("Enter Last Page Number",
                         min_value=0,
                         max_value=1000000000)


def get_col_widths(dataframe):
    # First we find the maximum length of the index column
    idx_max = max([len(str(s)) for s in dataframe.index.values] +
                  [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [idx_max] + [
        max([len(str(s)) for s in dataframe[col].values] + [len(col)])
        for col in dataframe.columns
    ]


def to_excel(df, text):
    global number
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    numberOfSheets = math.ceil(df.shape[0] / 12)
    for i in range(1, numberOfSheets + 1):
        number += 1
        sheetName = 'Sheet ' + str(i)
        df[(i - 1) * 12:i * 12].to_excel(writer,
                                         index=False,
                                         sheet_name=sheetName,
                                         startrow=1)
        stickers = pd.DataFrame()
        # define new dataFrame with 2 columns A and B
        stickers = pd.DataFrame(columns=['First Column', 'Second Column'])
        boxes = []
        for index, row in df[(i - 1) * 12:i * 12].iterrows():
            box = row["Type"].strip().upper() + "\n" + row["Color"].strip(
            ).title() + "\n" + row["Size"].strip().title()
            for i in range(1, int(row["Qty"]) + 1):
                boxes.append(box)
        count = 0
        for i in range(1, math.ceil(len(boxes) / 2) + 1):
            try:
                first = boxes[count]
            except:
                first = ""
            try:
                second = boxes[count + 1]
            except:
                second = ""
            stickers.loc[i] = [first, second]
            count += 1
            # count += 2
        # stickers.to_excel(writer,
        #                   index=False,
        #                   sheet_name=sheetName + " Stickers",
        #                   header=False)
        worksheet = writer.sheets[sheetName]
        # Stickers
        # stickersWorkSheet = writer.sheets[sheetName + " Stickers"]
        # stickersWorkSheet.set_default_row(height=71)
        format = workbook.add_format({
            "border": 0,
            "border_color": "black",
            "bold": True,
            "font_size": 14,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True
        })
        # stickersWorkSheet.set_column(0, 0, 50, format)
        # stickersWorkSheet.set_column(1, 1, 50, format)

        worksheet.merge_range(
            'A1:D1', f'Ambient - {text}',
            workbook.add_format({
                'bold': True,
                "border": 1,
                "border_color": "black",
                'font_size': 18,
                "border": 1,
                'align': 'center',
                'valign': 'vcenter'
            }))
        worksheet.merge_range(
            'E1:F1', f'{time.strftime("%d-%m-%Y", time.localtime())}',
            workbook.add_format({
                'bold': True,
                "border": 1,
                "border_color": "black",
                'font_size': 18,
                'align': 'center',
                'valign': 'vcenter'
            }))
        worksheet.merge_range(
            'G1:H1', f'Ambient-{number}',
            workbook.add_format({
                'bold': True,
                "border": 1,
                "border_color": "black",
                'font_size': 18,
                'align': 'center',
                'valign': 'vcenter'
            }))
        format = workbook.add_format({
            "border": 1,
            "border_color": "black",
            "bold": True,
            "font_size": 18,
            "align": "center",
            "valign": "vcenter",
        })
        worksheet.set_default_row(height=31.5)
        for i, width in enumerate(get_col_widths(df)):
            worksheet.set_column(i - 1, i - 1, width * 1.61 + 2, format)
    writer.close()
    processed_data = output.getvalue()
    return processed_data


def processTheSide(length):
    if "'" in length and '"' in length:
        foot = float(length.split("'")[0].strip())
        inch = float(length.split("'")[1].split("\"")[0].strip())
        length = foot + (inch / 12)
        return length
    if "'" in length and '"' not in length:
        foot = float(length.split("'")[0].strip())
        length = foot
        return length
    if "'" not in length and '"' in length:
        inch = float(length.split("\"")[0].strip())
        length = inch / 12
        return length
    return 100


def processLength(size):
    try:
        size = size.strip().upper()
        if "X" in size and "HEX" not in size:
            length = processTheSide(size.split("X")[0].strip())
            width = processTheSide(size.split("X")[1].strip())
        else:
            length = processTheSide(size.split(" ")[0].strip())
            width = processTheSide(size.split(" ")[0].strip())
        if length == 3:
            if width == 4 or width == 5:
                return "S"
        if length < 4 and width < 6:
            return "S"
        else:
            return "L"
    except:
        return "L"


col1, col2, col3 = st.columns(3)


with col1:
    st.header("Walmart")
    walmartOrderReport = st.file_uploader("Choose Walmart Order Report")
    if walmartOrderReport:
        walmartOrderReport = pd.read_excel(walmartOrderReport, engine="openpyxl", dtype={
            "Order#": str,
            "Qty": int,
            "SKU": str
        })

with col2:
    st.header("Wayfair")
    wayfairOrderReport = st.file_uploader("Choose Wayfair Order Report")
    if wayfairOrderReport:
        wayfairOrderReport = pd.read_excel(wayfairOrderReport, engine="openpyxl", dtype={
            "PO Number": str,
            "Quantity": int,
            "Item Number": str
        })

with col3:
    st.header("Amazon")
    amazonOrderReport = st.file_uploader("Choose Amazon Order Report")
    if amazonOrderReport:
        amazonOrderReport = pd.read_excel(amazonOrderReport, engine="openpyxl", dtype={
            "order-id": str,
            "sku": str,
            "quantity-purchased": int
        })

if st.button("Process"):
    with st.spinner("Getting the Job Done..."):
        allOrdersReport = []
        for index, row in amazonOrderReport.iterrows():
            allOrdersReport.append({
                "OrderID": row["order-id"],
                "SKU": row["sku"],
                "Qty": row["quantity-purchased"],
                "Source": "Amazon"
            })
        if walmartOrderReport is not None:
            for index, row in walmartOrderReport.iterrows():
                allOrdersReport.append({
                    "OrderID": row["Order#"],
                    "SKU": row["SKU"],
                    "Qty": row["Qty"],
                    "Source": "Walmart"
                })
        if wayfairOrderReport is not None:
            for index, row in wayfairOrderReport.iterrows():
                allOrdersReport.append({
                    "OrderID": row["PO Number"],
                    "SKU": row["Item Number"],
                    "Qty": row["Quantity"],
                    "Source": "Wayfair"
                })

        allOrdersReport = pd.DataFrame(allOrdersReport)

        # st.table(allOrdersReport)

        masterFile = pd.read_excel(
            "Ambient_Master_File.xlsx", engine="openpyxl")

        allOrdersReport["SKU"] = allOrdersReport["SKU"].str.upper()
        masterFile["WayfairMarchantSKU"] = masterFile["WayfairMarchantSKU"].str.upper()

        masterFile = masterFile.drop_duplicates(subset="WayfairMarchantSKU")
        masterFile.reset_index(drop=True, inplace=True)

        merged_df = allOrdersReport.merge(
            masterFile, left_on="SKU", right_on="WayfairMarchantSKU", how="left")

        merged_df["PRODUCT GROUP"].fillna("** UNKNOWN **", inplace=True)
        merged_df["COLOR"].fillna("** UNKNOWN **", inplace=True)
        merged_df["SIZE"].fillna("** UNKNOWN **", inplace=True)

        merged_df.rename(columns={"PRODUCT GROUP": "Type",
                                  "COLOR": "Color", "SIZE": "Size"}, inplace=True)

        # Create a new column "Single or Multiple" based on the logic
        merged_df['Single or Multiple'] = merged_df.groupby('OrderID')['Qty'].transform(
            lambda x: 'Multiple' if x.gt(1).any() else 'Single')

        # If a single OrderID comes multiple times, mark it as 'Multiple'
        merged_df.loc[merged_df['OrderID'].duplicated(
            keep=False), 'Single or Multiple'] = 'Multiple'

        # Sort the DataFrame by 'OrderID' to group them together
        merged_df.sort_values(by='OrderID', inplace=True)

        # Reset the index if needed
        merged_df.reset_index(drop=True, inplace=True)

        final_df = merged_df[["OrderID", "SKU",
                              "Qty", "Type", "Color", "Size", "Single or Multiple"]]

        final_df["Size"] = final_df["Size"].str.upper()
        final_df["Type"] = final_df["Type"].str.upper()

        final_df["Segmentation"] = [processLength(
            row["Size"]) for index, row in final_df.iterrows()]

        final_df["Cutter"] = ""
        final_df["Inspection"] = ""
        final_df["Label"] = ""

        # st.table(final_df)

        # Create a DataFrame for "Single" orders
        single_orders_df = final_df[final_df['Single or Multiple'] == 'Single']

        # Create a DataFrame for "Multiple" orders
        multiple_orders_df = final_df[final_df['Single or Multiple'] == 'Multiple']

        single_orders_small_sizes_list = single_orders_df[single_orders_df["Segmentation"] == "S"]

        single_orders_small_sizes_list_sorted = single_orders_small_sizes_list.sort_values(
            by=[
                "Type",
                "Color",
                "Size",
            ])

        single_orders_other_sizes_list = single_orders_df[single_orders_df["Segmentation"] != "S"]

        single_orders_other_sizes_list_neyland = single_orders_other_sizes_list[
            single_orders_other_sizes_list["Type"] == "NEYLAND"]

        single_orders_other_sizes_list = single_orders_other_sizes_list[
            single_orders_other_sizes_list["Type"] != "NEYLAND"]

        single_orders_other_sizes_list_sorted = single_orders_other_sizes_list.sort_values(
            by=[
                "Type",
                "Color",
                "Size",
            ])

        single_orders_other_sizes_list_neyland_sorted = single_orders_other_sizes_list_neyland.sort_values(
            by=[
                "Type",
                "Color",
                "Size",
            ])

        single_orders_small_sizes_list_sorted.drop(
            columns=["SKU", "Segmentation", "Single or Multiple"], inplace=True)
        single_orders_other_sizes_list_sorted.drop(
            columns=["SKU", "Segmentation", "Single or Multiple"], inplace=True)
        single_orders_other_sizes_list_neyland_sorted.drop(
            columns=["SKU", "Segmentation", "Single or Multiple"], inplace=True)
        multiple_orders_df.drop(
            columns=["SKU", "Segmentation", "Single or Multiple"], inplace=True)

        t = time.strftime("%d-%m-%Y %H:%M:%S", time.localtime())

        if single_orders_small_sizes_list_sorted.shape[0] > 0:
            st.markdown(
                f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(single_orders_small_sizes_list_sorted, "Single Orders Small Sizes")).decode()}" download="Ambient_Custom_Single_Orders_Small_Sizes_List_{t}.xlsx">✔️ Ambient Custom Single Orders Small Sizes List</a>',
                unsafe_allow_html=True)
        else:
            st.warning("No orders for Single Orders Small Sizes List.")
        if single_orders_other_sizes_list_sorted.shape[0] > 0:
            st.markdown(
                f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(single_orders_other_sizes_list_sorted, "Single Orders Other Sizes")).decode()}" download="Ambient_Custom_Single_Orders_Other_Sizes_List_{t}.xlsx">✔️ Ambient Custom Single Orders Other Sizes List</a>',
                unsafe_allow_html=True)
        else:
            st.warning("No orders for Single Orders Other Sizes List.")
        if single_orders_other_sizes_list_neyland_sorted.shape[0] > 0:
            st.markdown(
                f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(single_orders_other_sizes_list_neyland_sorted, "Single Orders Other Sizes Neyland Only")).decode()}" download="Ambient_Custom_Single_Orders_Other_Sizes_Neyland_Only_List_{t}.xlsx">✔️ Ambient Custom Single Orders Other Sizes Neyland Only List</a>',
                unsafe_allow_html=True)
        else:
            st.warning("No orders for Single Orders Other Sizes Neyland Only List.")
        if multiple_orders_df.shape[0] > 0:
            st.markdown(
                f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(multiple_orders_df, "Multi Orders")).decode()}" download="Ambient_Custom_Multi_Orders_List_{t}.xlsx">✔️ Ambient Custom Multi Orders List</a>',
                unsafe_allow_html=True)
        else:
            st.warning("No orders for Multi Orders List.")

        st.success(f"The last page number is {number}")
