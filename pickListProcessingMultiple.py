import streamlit as st
import pandas as pd
from io import BytesIO
import base64
import time

def get_col_widths(dataframe):
    # First we find the maximum length of the index column
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]


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
        worksheet.set_column(i-1, i-1, width + 2, format)
    writer.save()
    processed_data = output.getvalue()
    return processed_data


fmpMasterFile = pd.read_excel(
    "FMP_MASTER_DATA_FILE.xlsx",
    engine="openpyxl")

ambientMasterFile = pd.read_excel(
    "AMBIENT_MASTER_DATA_FILE.xlsx",
    engine="openpyxl")

fmpMasterFile = pd.concat([fmpMasterFile, ambientMasterFile], ignore_index=True)

st.write("""# FMP + Ambient Pick Lists Processing""")

fmpPicklist = st.file_uploader("Choose FMP Picklist")
ambientPicklist = st.file_uploader("Choose Ambient Picklist")

if st.button("Process Picklist"):
    with st.spinner("Getting the Job Done..."):
        fmpPickList = pd.read_excel(fmpPicklist,
                                    engine="openpyxl",
                                    dtype={
                                        "ProductID": str,
                                        "SingleItemOrderIDList": str,
                                        "MultiItemOrderIDList": str,
                                        "Size": str,
                                        "Color": str,
                                        "Product Group/Type": str,
                                        "Qty": int,
                                    })

        ambientPickList = pd.read_excel(ambientPicklist,
                                    engine="openpyxl",
                                    dtype={
                                        "ProductID": str,
                                        "SingleItemOrderIDList": str,
                                        "MultiItemOrderIDList": str,
                                        "Size": str,
                                        "Color": str,
                                        "Product Group/Type": str,
                                        "Qty": int,
                                    })

        fmpPickList["Company"] = "FMP"
        ambientPickList["Company"] = "Ambient"

        fmpPickList = pd.concat([fmpPickList, ambientPickList], ignore_index=True)

        fmpPickList = fmpPickList[[
            "SingleItemOrderIDList", "MultiItemOrderIDList", "ProductID",
            "Product Group/Type", "Qty", "Color", "Size", "Company"
        ]]

        #st.dataframe(fmpPickList)

        for index, row in fmpPickList.iterrows():
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

        fmpPickList = fmpPickList[[
            "SingleItemOrderIDList", "MultiItemOrderIDList",
            "Type", "Qty", "Color", "Size", "Cutter", "Inspection",
            "Label", "Company"
        ]]

        fmpPickList["Type"] = [
            row["Type"].upper() for index, row in fmpPickList.iterrows()
        ]
        fmpPickList["Size"] = [
            row["Size"].upper() for index, row in fmpPickList.iterrows()
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

        fmpCustomSingleOrdersList.rename(columns={"SingleItemOrderIDList": "OrderID"},
                                        inplace=True)
        fmpCustomMultiOrdersList.rename(columns={"MultiItemOrderIDList": "OrderID"},
                                        inplace=True)
        fmpCustomSingleOrdersList.drop(columns=["MultiItemOrderIDList"], inplace=True)
        fmpCustomMultiOrdersList.drop(columns=["SingleItemOrderIDList"], inplace=True)

        fmpCustomSingleOrdersSmallSizesList = fmpCustomSingleOrdersList.loc[
            fmpCustomSingleOrdersList["Size"].isin([
                "2' ROUND", "3' ROUND", "4' ROUND", "2' X 3'", "2' X 4'",
                '18" X 36" HALF ROUND', '20" X 40" HALF ROUND', "1.5' x 2.25'"
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
                '18" X 36" HALF ROUND', '20" X 40" HALF ROUND', "1.5' x 2.25'"
            ])]

        fmpCustomSingleOrdersOtherSizesListSorted = fmpCustomSingleOrdersOtherSizesList.sort_values(
            by=[
                "Type",
                "Color",
                "Size",
            ])

        fmpCustomMultiOrdersList = fmpCustomMultiOrdersList.sort_values(by=["OrderID"])

        # st.write(f"Cut Pieces ({fmpCutPiecesList.shape[0]})")
        # st.dataframe(fmpCutPiecesList)

        # st.write(f"Custom ({fmpCustomList.shape[0]})")
        # st.dataframe(fmpCustomList)

        # st.write(f"Custom Single Orders ({fmpCustomSingleOrdersList.shape[0]})")
        # st.dataframe(fmpCustomSingleOrdersList)

        # st.write(f"Custom Single Orders Small Sizes List ({fmpCustomSingleOrdersSmallSizesList.shape[0]})")
        # st.dataframe(fmpCustomSingleOrdersSmallSizesList)

        # st.write(f"Custom Single Orders Small Sizes List Sorted ({fmpCustomSingleOrdersSmallSizesListSorted.shape[0]})")
        # st.dataframe(fmpCustomSingleOrdersSmallSizesListSorted)

        # st.write(f"Custom Single Orders Other Sizes List ({fmpCustomSingleOrdersOtherSizesList.shape[0]})")
        # st.dataframe(fmpCustomSingleOrdersOtherSizesList)

        # st.write(f"Custom Single Orders Other Sizes List Sorted ({fmpCustomSingleOrdersOtherSizesListSorted.shape[0]})")
        # st.dataframe(fmpCustomSingleOrdersOtherSizesListSorted)

        # st.write(f"Custom Multi Orders ({fmpCustomMultiOrdersList.shape[0]})")
        # st.dataframe(fmpCustomMultiOrdersList)

    # st.download_button(label='✔️ Custom Single Orders Small Sizes List',
    #                 data=to_excel(fmpCustomSingleOrdersSmallSizesListSorted),
    #                 file_name='Custom_Single_Orders_Small_Sizes_List.xlsx')

    # st.download_button(label='✔️ Custom Single Orders Other Sizes List',
    #                 data=to_excel(fmpCustomSingleOrdersOtherSizesListSorted),
    #                 file_name='Custom_Single_Orders_Other_Sizes_List.xlsx')

    # st.download_button(label='✔️ Custom Multi Orders List',
    #                 data=to_excel(fmpCustomMultiOrdersList),
    #                 file_name='Custom_Multi_Orders_List.xlsx')

    # st.download_button(label='✔️ Cut Pieces List',
    #                 data=to_excel(fmpCutPiecesList),
    #                 file_name='Cut_Pieces_List.xlsx')

    t = time.strftime("%d-%m-%Y %H:%M:%S", time.localtime())

    st.markdown(f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(fmpCustomSingleOrdersSmallSizesListSorted)).decode()}" download="Custom_Single_Orders_Small_Sizes_List_{t}.xlsx">✔️ Custom Single Orders Small Sizes List</a>', unsafe_allow_html=True)
    st.markdown(
        f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(fmpCustomSingleOrdersOtherSizesListSorted)).decode()}" download="Custom_Single_Orders_Other_Sizes_List_{t}.xlsx">✔️ Custom Single Orders Other Sizes List</a>',
        unsafe_allow_html=True)
    st.markdown(
        f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(fmpCustomMultiOrdersList)).decode()}" download="Custom_Multi_Orders_List_{t}.xlsx">✔️ Custom Multi Orders List</a>',
        unsafe_allow_html=True)
    st.markdown(
        f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(fmpCutPiecesList)).decode()}" download="Cut_Pieces_List_{t}.xlsx">✔️ Cut Pieces List</a>',
        unsafe_allow_html=True)
