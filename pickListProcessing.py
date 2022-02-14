import streamlit as st
import pandas as pd
from io import BytesIO
import base64
import time

def get_col_widths(dataframe):
    # First we find the maximum length of the index column
    idx_max = max([len(str(s)) for s in dataframe.index.values] +
                  [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [idx_max] + [
        max([len(str(s)) for s in dataframe[col].values] + [len(col)])
        for col in dataframe.columns
    ]


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format = workbook.add_format({
        "border": 1,
        "border_color": "black",
        "bold": True,
        "align": "center",
        "valign": "vcenter",
    })
    worksheet.set_default_row(height=31.5)
    for i, width in enumerate(get_col_widths(df)):
        worksheet.set_column(i - 1, i - 1, width + 2, format)
    writer.save()
    processed_data = output.getvalue()
    return processed_data


fmpMasterFile = pd.read_excel(
    "FMP_MASTER_DATA_FILE.xlsx",
    engine="openpyxl")

st.write("""# FMP Pick List Processing""")
fmpPicklist = st.file_uploader("Choose FMP Picklist")

if st.button("Process Picklist"):
    with st.spinner("Getting the Job Done..."):
        fmpPickList = pd.read_excel(fmpPicklist,
                                    engine="openpyxl",
                                    dtype={
                                        "ProductID": str,
                                        "SingleItemOrderIDList": str,
                                        "MultiItemOrderIDList": str,
                                        "SingleOrderItemCount": int,
                                        "MultiOrderItemCount": int,
                                        "Size": str,
                                        "Color": str,
                                        "Product Group/Type": str,
                                        "Qty": int,
                                    })

        fmpPickList = fmpPickList[[
            "SingleItemOrderIDList", "MultiItemOrderIDList", "ProductID",
            "Product Group/Type", "Qty", "Color", "Size",
            "SingleOrderItemCount", "MultiOrderItemCount"
        ]]

        #st.dataframe(fmpPickList)

        fmpMasterFile["ProductID"] = fmpMasterFile["ProductID"].str.upper()
        fmpPickList["ProductID"] = fmpPickList["ProductID"].str.upper()

        for index, row in fmpPickList.iterrows():
            st.text(f"Processing {row['ProductID']}")
            fmpMasterFileRow = fmpMasterFile.loc[fmpMasterFile["ProductID"] ==
                                                row["ProductID"]]
            Type = fmpMasterFileRow["ProductGroupName"].values[0]
            Size = fmpMasterFileRow["SIZE"].values[0]
            Color = fmpMasterFileRow["COLOR"].values[0]
            if str(fmpPickList.loc[index, "Product Group/Type"]) == "nan":
                fmpPickList.loc[index, "Product Group/Type"] = Type
            if str(fmpPickList.loc[index, "Size"]) == "nan":
                fmpPickList.loc[index, "Size"] = Size
            if str(fmpPickList.loc[index, "Color"]) == "nan":
                fmpPickList.loc[index, "Color"] = Color

        #st.dataframe(fmpPickList)

        fmpPickList.drop(columns=["ProductID"], inplace=True)

        fmpPickList = fmpPickList.rename(columns={"Product Group/Type": "Type"})

        fmpPickList["Cutter"] = ""
        fmpPickList["Inspection"] = ""
        fmpPickList["Label"] = ""

        fmpPickList["Type"] = [
            str(row["Type"]).upper() for index, row in fmpPickList.iterrows()
        ]

        fmpPickList["Size"] = [
            str(row["Size"]).upper() for index, row in fmpPickList.iterrows()
        ]

        #st.dataframe(fmpPickList)

        fmpCutPiecesList = fmpPickList.loc[fmpPickList["Type"].isin(
            ["NYLON", "BCF", "FLORIDA"])]
        fmpCustomList = fmpPickList.loc[~fmpPickList["Type"].
                                        isin(["NYLON", "BCF", "FLORIDA"])]
        fmpCustomSingleOrdersList = fmpCustomList.loc[
            fmpCustomList["SingleItemOrderIDList"].notna()]
        fmpCustomMultiOrdersList = fmpCustomList.loc[
            fmpCustomList["MultiItemOrderIDList"].notna()]

        fmpCustomSingleOrdersList["Qty"] = fmpCustomSingleOrdersList[
            "SingleOrderItemCount"]
        fmpCustomMultiOrdersList["Qty"] = fmpCustomMultiOrdersList[
            "MultiOrderItemCount"]

        fmpCustomSingleOrdersList.rename(columns={"SingleItemOrderIDList": "OrderID"},
                                        inplace=True)
        fmpCustomMultiOrdersList.rename(columns={"MultiItemOrderIDList": "OrderID"},
                                        inplace=True)
        fmpCustomSingleOrdersList.drop(columns=["MultiItemOrderIDList"], inplace=True)
        fmpCustomMultiOrdersList.drop(columns=["SingleItemOrderIDList"], inplace=True)
        fmpCustomSingleOrdersList.drop(
            columns=["SingleOrderItemCount", "MultiOrderItemCount"],
            inplace=True)
        fmpCustomMultiOrdersList.drop(
            columns=["SingleOrderItemCount", "MultiOrderItemCount"],
            inplace=True)

        fmpCustomSingleOrdersSmallSizesList = fmpCustomSingleOrdersList.loc[
            fmpCustomSingleOrdersList["Size"].isin([
                "2' ROUND", "3' ROUND", "4' ROUND", "2' X 3'", "2' X 4'",
                '18" X 36" HALF ROUND', '20" X 40" HALF ROUND', "1.5' X 2.25'"
            ])]

        fmpCustomSingleOrdersSmallSizesListSorted = fmpCustomSingleOrdersSmallSizesList.sort_values(
            by=[
                "Type",
                "Color",
                "Size",
            ])

        fmpCustomSingleOrdersOtherSizesList = fmpCustomSingleOrdersList.loc[
            ~fmpCustomSingleOrdersList["Size"].isin([
                "2' ROUND", "3' ROUND", "4' ROUND", "2' X 3'", "2' X 4'",
                '18" X 36" HALF ROUND', '20" X 40" HALF ROUND', "1.5' X 2.25'"
            ])]

        fmpCustomSingleOrdersOtherSizesListSorted = fmpCustomSingleOrdersOtherSizesList.sort_values(
            by=[
                "Type",
                "Color",
                "Size",
            ])

        fmpCustomMultiOrdersList = fmpCustomMultiOrdersList.sort_values(by=["OrderID"])

    t = time.strftime("%d-%m-%Y %H:%M:%S", time.localtime())

    st.markdown(
        f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(fmpCustomSingleOrdersSmallSizesListSorted)).decode()}" download="Custom_Single_Orders_Small_Sizes_List_{t}.xlsx">✔️ Custom Single Orders Small Sizes List</a>',
        unsafe_allow_html=True)
    st.markdown(
        f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(fmpCustomSingleOrdersOtherSizesListSorted)).decode()}" download="Custom_Single_Orders_Other_Sizes_List_{t}.xlsx">✔️ Custom Single Orders Other Sizes List</a>',
        unsafe_allow_html=True)
    st.markdown(
        f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(fmpCustomMultiOrdersList)).decode()}" download="Custom_Multi_Orders_List_{t}.xlsx">✔️ Custom Multi Orders List</a>',
        unsafe_allow_html=True)
    st.markdown(
        f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(fmpCutPiecesList)).decode()}" download="Cut_Pieces_List_{t}.xlsx">✔️ Cut Pieces List</a>',
        unsafe_allow_html=True)
