import pandas as pd
import numpy as np
import requests
from datetime import datetime

# Download the Excel file from Google Sheets
url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQFvubpNHb1TQEPliUeeuyqWx30SLFagXt8CTDt1L4y4O_PLTSKiqQulEdnbNG-GdWKteAd7ueLB9f4/pub?output=xlsx"
response = requests.get(url)
with open("data.xlsx", "wb") as file:
    file.write(response.content)

# Load the Excel file and select the "Angus Steak House" sheet
data = pd.read_excel("data.xlsx", sheet_name="Angus Steak House")

# Promote headers
#data.columns = data.iloc[0]
# = data[1:]

# Change column types
data["Delivery Date 送货日期"] = pd.to_datetime(data["Delivery Date 送货日期"], errors="coerce")
data["Submission ID"] = data["Submission ID"].astype(str)

# Split 'My Products: Products' column by delimiter ")"
data["My Products: Products"] = data["My Products: Products"].str.split(")")
data = data.explode("My Products: Products")

# Further split columns and clean data
data[["Outlet 地址.1", "Outlet 地址.2", "Outlet 地址.3", "Outlet 地址.4", "Outlet 地址.5"]] = data["Outlet 地址"].str.split("-", expand=True)
data["My Products: Products"] = data["My Products: Products"].str.strip()
data[["My Products: Products.1", "My Products: Products.2"]] = data["My Products: Products"].str.split("-", n=1, expand=True)
data[["My Products: Products.2.1", "My Products: Products.2.2"]] = data["My Products: Products.2"].str.split("(", n=1, expand=True)
data[["My Products: Products.2.2.1", "My Products: Products.2.2.2", "My Products: Products.2.2.3"]] = data["My Products: Products.2.2"].str.split(", ", expand=True)

# Clean and rename columns
data["My Products: Products.2.2.1"] = data["My Products: Products.2.2.1"].str.replace("Amount: ", "").str.replace(" SGD", "").astype(float)
data["My Products: Products.2.2.2"] = data["My Products: Products.2.2.2"].str.replace("Quantity:", "").astype(float)
data["My Products: Products.2.2.3"] = data["My Products: Products.2.2.3"].str.replace(": ", "")
data = data.rename(columns={
    "Delivery Date 送货日期": "delivery_date_required",
    "Outlet 地址.1": "business_outlet",
    "Outlet 地址.2": "bill_to",
    "My Products: Products.1": "item_name",
    "My Products: Products.2.1": "supplier_item_code",
    "My Products: Products.2.2.1": "unit_price",
    "My Products: Products.2.2.2": "quantity_required",
    "My Products: Products.2.2.3": "uom",
    "Submission ID": "po_no",
    "Remark 注明": "remark",
})

# Add calculated columns
data["amount_required"] = data["unit_price"] * data["quantity_required"]
data["amount_supplier"] = data["amount_required"]
data["buyer_code"] = ""
data["order_date"] = pd.Timestamp.now()
data["purchase_order_date"] = data["order_date"]
data["specific_request"] = data["remark"]
data["po_no"] = data.apply(lambda row: row["po_no"] + "-F" if row["item_name"].startswith(("FR", "ZF", "FSI")) else row["po_no"] + "-D", axis=1)
data["type"] = data["po_no"].apply(lambda x: 1019 if x.endswith("-F") else 1016)
data["Supplier"] = "Lim Siang Huat Pte Ltd"

# Reorder columns
column_order = [
    "po_no", "delivery_date_required", "business_outlet", "bill_to", "item_name", "supplier_item_code",
    "unit_price", "quantity_required", "uom", "amount_required", "buyer_code", "Supplier", "order_date",
    "quantity_supplier", "weight", "amount_supplier", "delivery_date_supplier", "specific_request",
    "purchase_order_date", "remark", "type"
]
data = data.reindex(columns=column_order, fill_value="")

# Save the transformed data to a new Excel file
data.to_excel("transformed_data.xlsx", index=False)
