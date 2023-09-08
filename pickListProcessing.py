import streamlit as st
import pandas as pd
from io import BytesIO
import base64
import time
import math


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
    dfs = [df[(i - 1) * 12:i * 12] for i in range(1, numberOfSheets + 2)]
    xo = 0
    for curr, next in zip(dfs, dfs[1:]):
        number += 1
        xo += 1
        sheetName = 'Sheet ' + str(xo)
        last = curr.tail(1).get("OrderID").str.split(",").tolist()
        first = next.head(1).get("OrderID").str.split(",").tolist()
        if len(last) > 0 and len(first) > 0:
            last = set(last[0])
            first = set(first[0])
            while len(last.intersection(first)) > 0:
                first_row = next.iloc[0]
                next = next.iloc[1:]
                dfs[xo] = dfs[xo].iloc[1:]
                curr = curr.append(first_row, ignore_index=True)
                last = curr.tail(1).get("OrderID").str.split(",").tolist()
                first = next.head(1).get("OrderID").str.split(",").tolist()
                if len(last) > 0 and len(first) > 0:
                    last = set(last[0])
                    first = set(first[0])
                else:
                    last = set()
                    first = set()
        if curr.shape[0] == 0:
            continue
        curr.to_excel(writer,
                      index=False,
                      sheet_name=sheetName,
                      startrow=1)
        worksheet = writer.sheets[sheetName]
        format = workbook.add_format({
            "border": 0,
            "border_color": "black",
            "bold": True,
            "font_size": 14,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True
        })

        worksheet.merge_range(
            'A1:D1', f'FMP - {text}',
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
            'G1:H1', f'F{number}',
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


st.write("""# FMP Pick List Processing""")

try:
    f = open("fmpListLastPageNumber.txt", "r")
    number = int(f.read().strip())
    f.close()
except:
    number = 0

number = st.number_input("Enter Last Page Number", value=number, min_value=0, max_value=1000000000)

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

        fmpMasterFile = pd.read_excel(
            "FMP_MASTER_DATA_FILE_17_10_2022.xlsx", engine="openpyxl")

        fmpPickList = fmpPickList[[
            "SingleItemOrderIDList", "MultiItemOrderIDList", "ProductID",
            "Product Group/Type", "Qty", "Color", "Size",
            "SingleOrderItemCount", "MultiOrderItemCount"
        ]]

        # st.dataframe(fmpPickList)

        fmpMasterFile["ProductID"] = fmpMasterFile["ProductID"].str.upper()
        fmpPickList["ProductID"] = fmpPickList["ProductID"].str.upper()

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
                        return "L"
                if length < 4 and width < 6:
                    return "S"
                else:
                    return "L"
            except:
                return "L"

        fmpPickList["SizeType"] = [
            processLength(row["Size"])
            for index, row in fmpPickList.iterrows()
        ]

        validCustoms = fmpMasterFile["ProductGroupName"].dropna(
        ).unique().tolist()
        validCustoms = [x.upper() for x in validCustoms]

        for index, row in fmpPickList.iterrows():
            st.text(f"Processing {row['ProductID']}")
            fmpMasterFileRow = fmpMasterFile.loc[fmpMasterFile["ProductID"] ==
                                                 row["ProductID"]]
            thingsNotFound = []
            thereWasErrorForThisSKU = False
            try:
                Type = fmpMasterFileRow["ProductGroupName"].values[0]
            except:
                thereWasErrorForThisSKU = True
                thingsNotFound.append("Type")
                Type = "UNKNOWN Type"
            try:
                Size = fmpMasterFileRow["SIZE"].values[0]
            except:
                thereWasErrorForThisSKU = True
                thingsNotFound.append("Size")
                Size = "Unknown Size"
            try:
                Color = fmpMasterFileRow["COLOR"].values[0]
            except:
                thereWasErrorForThisSKU = True
                thingsNotFound.append("Color")
                if "RUG PAD" in row["ProductID"].upper():
                    Color = "RUG PAD"
                    st.warning(f"RUG PAD added in Color Field for {row['ProductID']}")
                else:
                    Color = "Unknown Color"
            itemsFound = []
            if str(fmpPickList.loc[index, "Product Group/Type"]) == "nan":
                fmpPickList.loc[index, "Product Group/Type"] = Type
            else:
                itemsFound.append("Type")
            if str(fmpPickList.loc[index, "Size"]) == "nan":
                fmpPickList.loc[index, "Size"] = Size
            else:
                itemsFound.append("Size")
            if str(fmpPickList.loc[index, "Color"]) == "nan":
                fmpPickList.loc[index, "Color"] = Color
            else:
                itemsFound.append("Color")

            if all(custom not in row["Product Group/Type"] for custom in validCustoms):
                st.warning(
                    f'Possible Non-Custom order, so skipping: {row["ProductID"]}')
                fmpPickList.drop(index, inplace=True)
                continue
            
            if thereWasErrorForThisSKU:
                st.warning(
                    f"Error Processing {row['ProductID']} | Items not found in MASTER File: {thingsNotFound}")
                try:
                    if len(itemsFound) > 0:
                        st.success(f'But these items were found in Picklist and were filled accordingly: {itemsFound}')
                        if len(set(thingsNotFound).intersection(itemsFound)) == 0:
                            st.error(f'One or more Items were not found anywhere, neither in master File nor in Picklist: {",".join(list(set(thingsNotFound).intersection(itemsFound)))}')
                    else:
                        st.error(
                            f'Please check manually as any Items were not even found in the Picklist.', icon="🚨")
                except:
                    pass

        # st.dataframe(fmpPickList)

        fmpPickList.drop(columns=["ProductID"], inplace=True)

        fmpPickList = fmpPickList.rename(
            columns={"Product Group/Type": "Type"})

        fmpPickList["Cutter"] = ""
        fmpPickList["Inspection"] = ""
        fmpPickList["Label"] = ""

        fmpPickList["Type"] = [
            str(row["Type"]).upper() for index, row in fmpPickList.iterrows()
        ]

        fmpPickList["Size"] = [
            str(row["Size"]).upper() for index, row in fmpPickList.iterrows()
        ]

        # st.dataframe(fmpPickList)

        fmpCutPiecesList = fmpPickList.loc[fmpPickList["Type"].isin(
            ["NYLON", "BCF", "FLORIDA"])]
        fmpCustomList = fmpPickList.loc[~fmpPickList["Type"].isin([
            "NYLON", "BCF", "FLORIDA", "UTTERMOST", "BUTLER", "COLONIAL MILL",
            "RADICI", "UNITED WEAVER"
        ])]
        fmpCustomSingleOrdersList = fmpCustomList.loc[
            fmpCustomList["SingleItemOrderIDList"].notna()]
        fmpCustomMultiOrdersList = fmpCustomList.loc[
            fmpCustomList["MultiItemOrderIDList"].notna()]

        fmpCustomSingleOrdersList["Qty"] = fmpCustomSingleOrdersList[
            "SingleOrderItemCount"]
        fmpCustomMultiOrdersList["Qty"] = fmpCustomMultiOrdersList[
            "MultiOrderItemCount"]

        fmpCustomSingleOrdersList.rename(
            columns={"SingleItemOrderIDList": "OrderID"}, inplace=True)
        fmpCustomMultiOrdersList.rename(
            columns={"MultiItemOrderIDList": "OrderID"}, inplace=True)
        fmpCustomSingleOrdersList.drop(columns=["MultiItemOrderIDList"],
                                       inplace=True)
        fmpCustomMultiOrdersList.drop(columns=["SingleItemOrderIDList"],
                                      inplace=True)
        fmpCustomSingleOrdersList.drop(
            columns=["SingleOrderItemCount", "MultiOrderItemCount"],
            inplace=True)
        fmpCustomMultiOrdersList.drop(
            columns=["SingleOrderItemCount", "MultiOrderItemCount"],
            inplace=True)

        fmpCustomSingleOrdersSmallSizesList = fmpCustomSingleOrdersList.loc[
            fmpCustomSingleOrdersList["SizeType"].isin(["S"])]

        fmpCustomSingleOrdersSmallSizesListSorted = fmpCustomSingleOrdersSmallSizesList.sort_values(
            by=[
                "Type",
                "Color",
                "Size",
            ])

        fmpCustomSingleOrdersOtherSizesList = fmpCustomSingleOrdersList.loc[
            ~fmpCustomSingleOrdersList["SizeType"].isin(["S"])]

        fmpCustomSingleOrdersOtherSizesListNeyLand = fmpCustomSingleOrdersOtherSizesList.loc[
            fmpCustomSingleOrdersOtherSizesList["Type"].isin(["NEYLAND"])]

        fmpCustomSingleOrdersOtherSizesList = fmpCustomSingleOrdersOtherSizesList.loc[
            ~fmpCustomSingleOrdersOtherSizesList["Type"].isin(["NEYLAND"])]

        fmpCustomSingleOrdersOtherSizesListSorted = fmpCustomSingleOrdersOtherSizesList.sort_values(
            by=[
                "Type",
                "Color",
                "Size",
            ])

        fmpCustomSingleOrdersOtherSizesListNeyLandSorted = fmpCustomSingleOrdersOtherSizesListNeyLand.sort_values(
            by=[
                "Type",
                "Color",
                "Size",
            ])

        fmpCustomMultiOrdersList = fmpCustomMultiOrdersList.sort_values(
            by=["OrderID"])

    fmpCustomSingleOrdersSmallSizesListSorted.drop(columns=["SizeType"],
                                                   inplace=True)
    fmpCustomSingleOrdersOtherSizesListSorted.drop(columns=["SizeType"],
                                                   inplace=True)
    fmpCustomSingleOrdersOtherSizesListNeyLandSorted.drop(columns=["SizeType"],
                                                          inplace=True)
    fmpCustomMultiOrdersList.drop(columns=["SizeType"], inplace=True)

    t = time.strftime("%d-%m-%Y %H:%M:%S", time.localtime())

    if fmpCustomSingleOrdersSmallSizesListSorted.shape[0] > 0:
        st.markdown(
            f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(fmpCustomSingleOrdersSmallSizesListSorted, "Single Orders Small Sizes")).decode()}" download="Custom_Single_Orders_Small_Sizes_List_{t}.xlsx">✔️ Custom Single Orders Small Sizes List</a>',
            unsafe_allow_html=True)
    else:
        st.warning("No Orders for Single Orders Small Sizes.")
    if fmpCustomSingleOrdersOtherSizesListSorted.shape[0] > 0:
        st.markdown(
            f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(fmpCustomSingleOrdersOtherSizesListSorted, "Single Orders Other Sizes")).decode()}" download="Custom_Single_Orders_Other_Sizes_List_{t}.xlsx">✔️ Custom Single Orders Other Sizes List</a>',
            unsafe_allow_html=True)
    else:
        st.warning("No Orders for Single Orders Other Sizes.")
    if fmpCustomSingleOrdersOtherSizesListNeyLandSorted.shape[0] > 0:
        st.markdown(
            f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(fmpCustomSingleOrdersOtherSizesListNeyLandSorted, "Single Orders Other Sizes Neyland Only")).decode()}" download="Custom_Single_Orders_Other_Sizes_Neyland_Only_List_{t}.xlsx">✔️ Custom Single Orders Other Sizes Neyland Only List</a>',
            unsafe_allow_html=True)
    else:
        st.warning("No Orders for Single Orders Other Sizes Neyland Only.")
    if fmpCustomMultiOrdersList.shape[0] > 0:
        st.markdown(
            f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(fmpCustomMultiOrdersList, "Multi Orders")).decode()}" download="Custom_Multi_Orders_List_{t}.xlsx">✔️ Custom Multi Orders List</a>',
            unsafe_allow_html=True)
    else:
        st.warning("No Orders for Multi Orders.")
    st.success(f"The last sheet number is {number}.")

    f = open("fmpListLastPageNumber.txt", "w")
    f.write(str(number))
    f.close()
    # st.markdown(
    #     f'<a href="data:application/octet-stream;base64,{base64.b64encode(to_excel(fmpCutPiecesList, "Cut Pieces")).decode()}" download="Cut_Pieces_List_{t}.xlsx">✔️ Cut Pieces List</a>',
    #     unsafe_allow_html=True)
