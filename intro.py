# import streamlit as st
# import pandas as pd
# from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
# from io import BytesIO

# # JavaScript to toggle cell value between "P" and "A" on click, with restrictions
# cell_click_js = JsCode("""
# function(e) {
#     let cell = e.api.getFocusedCell();
#     let rowIndex = cell.rowIndex;
#     let colId = cell.column.colId;
#     let colIndex = cell.column.instanceId;
#     // Allow toggling only for row index >= 1 and column index >= 8 and colIndex <= 67
#     if (rowIndex >= 1 && colIndex >= 8 && colIndex <= 67) {
#         let rowNode = e.api.getRowNode(rowIndex);
#         let cellValue = rowNode.data[colId];
#         if (cellValue === "P") {
#             rowNode.setDataValue(colId, "A");
#         } else {
#             rowNode.setDataValue(colId, "P");
#         }
#         // Update the 'TOTAL' column
#         let newRowData = rowNode.data;
#         newRowData['TOTAL'] = Object.values(newRowData).slice(8, 68).filter(value => value === "P").length;
#         rowNode.setData(newRowData);
#     }
# }
# """)

# # File uploader
# uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
# st.markdown(" ")

# if uploaded_file:
#     # Load the uploaded Excel file into a DataFrame
#     content = uploaded_file.read()
#     df = pd.read_excel(BytesIO(content), engine='openpyxl')
    
#     # Initialize 'TOTAL' column
#     df['TOTAL'] = df.iloc[:, 8:68].apply(lambda row: (row == 'P').sum(), axis=1)
    
#     # Set up AgGrid options
#     gb = GridOptionsBuilder.from_dataframe(df)

#     # Configure default column properties
#     gb.configure_default_column(editable=True)

#     # Set columns <= 8 and > 67 as non-editable
#     for col in df.columns[:9]:
#         gb.configure_column(col, editable=False)

#     for col in df.columns[68:]:
#         gb.configure_column(col, editable=False)

#     # Set 'TOTAL' column as non-editable
#     gb.configure_column('TOTAL', editable=False)

#     gb.configure_grid_options(onCellClicked=cell_click_js)
#     grid_options = gb.build()

#     # Display the DataFrame with AgGrid
#     grid_response = AgGrid(
#         df,
#         gridOptions=grid_options,
#         editable=True,
#         height=500,
#         allow_unsafe_jscode=True,  # Set it to True to allow custom JsCode
#         reload_data=True
#     )

#     # Updated DataFrame after clicking and editing
#     updated_df = pd.DataFrame(grid_response['data'])

#     # Save changes button
#     if st.button('SAVE CHANGES'):
#         # Create a new Excel workbook in memory
#         output = BytesIO()
#         writer = pd.ExcelWriter(output, engine='openpyxl')
#         updated_df.to_excel(writer, index=False, sheet_name='Sheet1')
#         writer.close()
#         output.seek(0)

#         # Display updated DataFrame
#         st.write("Updated DataFrame:")
#         st.write(updated_df)

#         # Provide a download link for the new Excel file
#         st.download_button(
#             label='Download Updated Excel File',
#             data=output,
#             file_name='updated_data.xlsx',
#             mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
#         )
        
#         st.success("Changes successfully saved. Download the updated file above.")

import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
from io import BytesIO
from datetime import datetime

# JavaScript to toggle cell value between "P" and "A" on click, with restrictions
cell_click_js = JsCode("""
function(e) {
    let cell = e.api.getFocusedCell();
    let rowIndex = cell.rowIndex;
    let colId = cell.column.colId;
    let colIndex = cell.column.instanceId;
    // Allow toggling only for row index >= 1 and column index >= 8 and colIndex <= 67
    if (rowIndex >= 1 && colIndex >= 8 && colIndex <= 67) {
        let rowNode = e.api.getRowNode(rowIndex);
        let cellValue = rowNode.data[colId];
        if (cellValue === "P") {
            rowNode.setDataValue(colId, "A");
        } else {
            rowNode.setDataValue(colId, "P");
        }
        // Update the 'TOTAL' column
        let newRowData = rowNode.data;
        newRowData['TOTAL'] = Object.values(newRowData).slice(8, 68).filter(value => value === "P").length;
        rowNode.setData(newRowData);
    }
}
""")

# File uploader
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
st.markdown(" ")

if uploaded_file:
    # Load the uploaded Excel file into a DataFrame
    content = uploaded_file.read()
    df = pd.read_excel(BytesIO(content), engine='openpyxl')
    
    # Initialize 'TOTAL' column
    df['TOTAL'] = df.iloc[:, 8:68].apply(lambda row: (row == 'P').sum(), axis=1)
    
    # Set up AgGrid options
    gb = GridOptionsBuilder.from_dataframe(df)

    # Configure default column properties
    gb.configure_default_column(editable=True)

    # Set columns <= 8 and > 67 as non-editable
    for col in df.columns[:9]:
        gb.configure_column(col, editable=False)

    for col in df.columns[68:]:
        gb.configure_column(col, editable=False)

    # Set 'TOTAL' column as non-editable
    gb.configure_column('TOTAL', editable=False)

    gb.configure_grid_options(onCellClicked=cell_click_js)
    grid_options = gb.build()

    # Display the DataFrame with AgGrid
    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        editable=True,
        height=500,
        allow_unsafe_jscode=True,  # Set it to True to allow custom JsCode
        reload_data=True
    )

    # Updated DataFrame after clicking and editing
    updated_df = pd.DataFrame(grid_response['data'])

    # Shaka dropdown and date picker
    shaka_option = st.selectbox("SHAKHA", ["SAI MANDIR","JESHTH NAGARIK MAANCH","GANESH MAIDAN","PARMARTH NIKETAN","SANKALP SIDDHI","BEST COLONY","DATTA MANDIR","MANGALMURTI GANESH MANDIR","PANDUMASTER"])
    selected_date = st.date_input("Select Date", datetime.today())

    # Save changes button
    if st.button('SAVE CHANGES'):
        # Create a new Excel workbook in memory
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='openpyxl')
        updated_df.to_excel(writer, index=False, sheet_name='Sheet1')
        writer.close()
        output.seek(0)

        # Generate file name based on dropdown and date picker
        file_name = f"{shaka_option}_{selected_date.strftime('%Y-%m-%d')}.xlsx"

        # Display updated DataFrame
        st.write("Updated DataFrame:")
        st.write(updated_df)

        # Provide a download link for the new Excel file
        st.download_button(
            label='Download Updated Excel File',
            data=output,
            file_name=file_name,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
        st.success(f"Changes successfully saved. Download the updated file as {file_name} above.")
