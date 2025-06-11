# Automatic Report Cleaner & Generator

A Streamlit web application that allows users to upload or use a sample sales dataset, automatically clean it, and generate a downloadable Excel report with visual summaries.

---

## Features

- Supports uploading your own Excel file (.xlsx)
- Automatically cleans erroneous or missing values
- Generates summary reports by Region, Product, and Shipment Mode
- Embeds Excel charts for visual insights
- Offers one-click download of cleaned data and final report
- Warns users about expected file structure to ensure smooth processing

---

## Expected Excel Columns

The input Excel file must contain the following case-sensitive columns:

- Order ID
- Order Date  
- Ship Date  
- Region  
- Product ID  
- Sales  
- Profit  
- Quantity  
- Ship Mode  

---
