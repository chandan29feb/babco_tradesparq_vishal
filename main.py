import streamlit as st
import pandas as pd
import numpy as np
import io
from rapidfuzz import fuzz, process

st.set_page_config(page_title="Container Analysis Tool", layout="wide")
st.title("Container Analysis Tool for Tradesparq Exports")

uploaded_files = st.file_uploader(
    "Upload all Excel files", 
    accept_multiple_files=True, 
    type=["xlsx", "xls"]
)

required_columns = [
    "Importer", "Date", "Master Bill Number", "Quantity", 
    "Value(USD)", "Unit Price(USD)", "Description"
]

ignore_columns_in_cleaned_sheet = [
    'HS Code Description','Importer Address','Importer Contact','Exporter Address','Exporter Contact' 'Packaging type',
    'Number of packages','Package unit','TEU','Freight fee','Insurance fee',
    'Loading Place','Unloading Place','Customs','incoterms','Carrier','VOCC',
    'Vessel Name','Voyage','House Bill Number','Customs Declaration Number'
]

def normalize_importer_names(df, similarity_threshold=90):
    df['Normalized_Importer'] = df['Importer'].astype(str).str.upper().str.replace(r'[^A-Z0-9 ]', '', regex=True).str.strip()

    unique_names = []
    name_mapping = {}

    for name in df['Normalized_Importer'].unique():
        if not unique_names:
            unique_names.append(name)
            name_mapping[name] = name
            continue

        match, score, _ = process.extractOne(name, unique_names, scorer=fuzz.token_sort_ratio)

        if score >= similarity_threshold:
            name_mapping[name] = match
        else:
            unique_names.append(name)
            name_mapping[name] = name

    return df

if uploaded_files:
    all_data = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, header=1)
            df = df.iloc[:, 1:]

            if df.empty or df.dropna(how='all').shape[0] == 0:
                st.warning(f"File '{file.name}' is empty and was skipped.")
                continue

            df.columns = [col.strip() for col in df.columns]
            missing = [col for col in required_columns if col not in df.columns]
            if missing:
                st.warning(f"File '{file.name}' is missing required columns: {missing}")
                continue

            df['source_file'] = file.name
            all_data.append(df)
        except Exception as e:
            st.error(f"Error reading {file.name}: {e}")

    if not all_data:
        st.stop()

    df = pd.concat(all_data, ignore_index=True)
    df['Importer'] = df['Importer'].astype(str).str.strip().str.upper()
    df = normalize_importer_names(df)
    
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    
    df['Date_str'] = df['Date'].dt.strftime('%B%d')
    df['Master Bill Number'] = df['Master Bill Number'].astype(str).replace("nan", np.nan)
    df['Master Bill Number'] = df['Master Bill Number'].fillna(df['Importer'] + " " + df['Date_str'])

    df['Unique Master Bill Number'] = df['Master Bill Number'].astype(str).str.strip()
    df['Container Name'] = df['Unique Master Bill Number']

    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')
    df['Weight (kgs)'] = df['Quantity']
    df['Value (USD)'] = pd.to_numeric(df['Value(USD)'], errors='coerce')
    df['Unit Price (USD)'] = pd.to_numeric(df['Unit Price(USD)'], errors='coerce')
    df['Shipment Cost'] = df['Value (USD)']

    products_per_container = (
        df.groupby('Container Name')['Description']
        .apply(lambda x: list(x.dropna().unique()))
        .reset_index(name="Products List")
    )
    products_per_container.insert(
    loc=1,
    column='Total Products in Container',
    value=products_per_container['Products List'].apply(len)
    )

    weight_per_product = (
        df.groupby(['Container Name', 'Description'])['Weight (kgs)']
        .sum().reset_index()
    )

    shipment_cost_per_container = (
        df.groupby('Container Name')['Shipment Cost']
        .sum().reset_index(name='Total Shipment Cost (USD)')
    )

    revenue_per_importer = (
        df.groupby('Importer')['Value (USD)']
        .sum().reset_index(name='Total Value(USD) per Importer')
    )

    main_output = io.BytesIO()
    with pd.ExcelWriter(main_output, engine='xlsxwriter') as writer:

        def write_sheet(df_sheet, sheet_name, drop_columns=None):
            if drop_columns:
                df_sheet = df_sheet.drop(columns=[col for col in drop_columns if col in df_sheet.columns])
                
            sort_col = next(
            (
                col for col in df_sheet.columns 
                if any(k in col.lower() for k in ['cost', 'revenue', 'weight', 'value']) 
                and pd.api.types.is_numeric_dtype(df_sheet[col])
            ),
            None
            )
            
            if sort_col:
                df_sheet = df_sheet.sort_values(by=sort_col, ascending=False)

            df_sheet.to_excel(writer, index=False, sheet_name=sheet_name)
            worksheet = writer.sheets[sheet_name]

            worksheet.freeze_panes(1, 0)

            header_format = writer.book.add_format({'bold': True})
            standard_number_format = writer.book.add_format({'num_format': '#,##,##0'})
            
            for col_num, value in enumerate(df_sheet.columns.values):
                worksheet.write(0, col_num, value, header_format)

                max_len = max(
                    df_sheet[value].astype(str).map(len).max() if not df_sheet[value].isnull().all() else 10,
                    len(value)
                ) + 5
                
                col_lower = value.lower()
                
                if pd.api.types.is_datetime64_any_dtype(df_sheet[value]):
                    date_format = writer.book.add_format({'num_format': 'yyyy-mm-dd'})
                    worksheet.set_column(col_num, col_num, max_len + 5, date_format)
                elif any(keyword in col_lower for keyword in ['cost', 'revenue', 'weight', 'value']) and \
                    pd.api.types.is_numeric_dtype(df_sheet[value]):
                    worksheet.set_column(col_num, col_num, max_len, standard_number_format)
                else:
                    worksheet.set_column(col_num, col_num, max_len)

        write_sheet(df.copy(), "Cleaned Data", drop_columns=ignore_columns_in_cleaned_sheet)
        write_sheet(products_per_container, "Products per Container")
        write_sheet(weight_per_product, "Weight per Product")
        write_sheet(shipment_cost_per_container, "Shipment Cost")
        write_sheet(revenue_per_importer, "Total Value per Importer")
        
    st.success("Analysis complete. Download your report below:")
    st.download_button(
        label="Download Container Analysis Excel",
        data=main_output.getvalue(),
        file_name="Container_Analysis_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
