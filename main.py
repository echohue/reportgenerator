import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt
from dateparser import parse
from io import BytesIO
import xlsxwriter
import re
st.set_page_config(page_title="Excel Report Generator", layout="centered")

st.title("Automatic Report Cleaner & Generator")


file_path = 'sales_data_with_errors_region.xlsx'
expected_columns = [
    'Order ID','Order Date', 'Ship Date', 'Region',
    'Product ID', 'Sales', 'Profit', 'Quantity', 'Ship Mode'
]

# 🔔 Expected structure notice in an expander
with st.expander("⚠️ View Required Column Structure", expanded=True):
    st.markdown(
        """
        Please ensure your Excel file includes the following columns **(case-sensitive)**:
        """
    )
    st.markdown(
        "<ul>" + "".join(f"<li><code>{col}</code></li>" for col in expected_columns) + "</ul>",
        unsafe_allow_html=True
    )

st.markdown("### Choose How to Load Data")










option = st.radio("Choose how to load data:", ["Use sample file", "Upload your own file"])

if option == "Use sample file":
    st.info("Using the built-in sample file.")
    df = pd.read_excel(file_path)
elif option == "Upload your own file":
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        missing_cols = [col for col in expected_columns if col not in df.columns]
        if missing_cols:
            st.error(f"❌ The following required columns are missing: {', '.join(missing_cols)}")
            st.stop()
        else:
            st.success("✅ File matches the required structure.")
    else:
        st.warning("Please upload a valid Excel file.")
        st.stop()


def date_parser(x):
    if bool(re.search("^\d\d/\d\d/\d\d\d\d$",x)):
        return dt.datetime.strptime(x,"%d/%m/%Y")
    return parse(x)


def data_cleaning(df):
    df.drop_duplicates(inplace=True)
    df['Order Date'] = df['Order Date'].apply(date_parser)
    df['Ship Date'] = df['Ship Date'].apply(date_parser)
    df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
    df = df.apply(lambda x: x.replace({np.NaN:"Unknown"}) if x.dtype == "object" else x)
    df = df.apply(lambda x:x.replace({np.NaN:0}) if x.dtype == "int64" or x.dtype == "float64" else x)
    return df

df_cleaned = data_cleaning(df)
df_cleaned.to_excel('report_cleaned.xlsx',index=False)

with open("report_cleaned.xlsx", "rb") as file:
    st.download_button(
        label="Download Cleaned Report",
        data=file,
        file_name="report_cleaned.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

cols_to_group = ['Sales','Profit','Quantity']
summary_region = df_cleaned.groupby('Region')[cols_to_group].sum().reset_index()
summary_product = df_cleaned.groupby('Product ID')[cols_to_group].sum().reset_index()
df_cleaned['Delivery Time(Days)'] = (df_cleaned['Ship Date']-df_cleaned['Order Date']).dt.days
deliver_shipmode = df_cleaned.groupby('Ship Mode')['Delivery Time(Days)'].mean().reset_index()

def report_generator(df_cleaned):
    output = BytesIO()
    with pd.ExcelWriter(output,engine='xlsxwriter') as writer:
        df_cleaned.to_excel(writer,sheet_name = 'cleaned_report',index=False)
        summary_region.to_excel(writer,sheet_name = 'summary_region',index=False)
        summary_product.to_excel(writer,sheet_name = 'summary_product', index=False)
        deliver_shipmode.to_excel(writer,sheet_name= 'Shipment_delivery_analysis', index = False)
        workbook = writer.book
        worksheet_region = writer.sheets['summary_region']
        chart_region = workbook.add_chart({'type':'column'})
        chart_region.add_series({
            'name': 'Sales by Region',
            'categories': ['summary_region', 1, 0, len(summary_region), 0],
            'values': ['summary_region', 1, 1, len(summary_region), 1],
        })
        chart_region.add_series({
            'name': 'Profit by Region',
            'categories': ['summary_region', 1, 0, len(summary_region), 0],
            'values': ['summary_region', 1, 3, len(summary_region), 3],
        })
        chart_region.add_series({
            'name': 'Quantity by Region',
            'categories': ['summary_region', 1, 0, len(summary_region), 0],
            'values': ['summary_region', 1, 2, len(summary_region), 2],
        })
        worksheet_region.insert_chart('G7', chart_region)

        worksheet_product = writer.sheets['summary_product']
        chart_product = workbook.add_chart({'type': 'column'})

        chart_product.add_series({
            'name': 'Sales by Product',
            'categories': ['summary_product', 1, 0, len(summary_product), 0],
            'values': ['summary_product', 1, 1, len(summary_product), 1],
        })
        chart_product.add_series({
            'name': 'Profit by Product',
            'categories': ['summary_product', 1, 0, len(summary_product), 0],
            'values': ['summary_product', 1, 3, len(summary_product), 3],
        })
        chart_product.add_series({
            'name': 'Quantity by Product',
            'categories': ['summary_product', 1, 0, len(summary_product), 0],
            'values': ['summary_product', 1, 2, len(summary_product), 2],
        })
        worksheet_product.insert_chart('G7', chart_product)

        worksheet_shipment = writer.sheets['Shipment_delivery_analysis']
        chart_del = workbook.add_chart({'type': 'column'})
        chart_del.add_series({
            'name': 'Delivery Time by Shipment Mode',
            'categories': ['Shipment_delivery_analysis', 1, 0, len(deliver_shipmode), 0],
            'values': ['Shipment_delivery_analysis', 1, 1, len(deliver_shipmode), 1],
        })
        worksheet_shipment.insert_chart('G7', chart_del)
    output.seek(0)
    return output
excel_file = report_generator(df_cleaned)
st.download_button(
    label="Download Excel Report(Report Summary and Graph)",
    data=excel_file,
    file_name="report_summary.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)