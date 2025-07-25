import pandas as pd
import numpy as np
import requests
from datetime import datetime
import os
import subprocess
import logging
from tqdm import tqdm
import time
from pathlib import Path
import shutil
import cx_Oracle as oracledb
import cx_Oracle
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import win32com.client
import sys




path = r"C:\Simplr\WhatsAPP_simplr\Import\PO_ZG.xlsx"
archive = r"C:\Simplr\WhatsAPP_simplr\Archive\PO_ZG.xlsx"

sender_email = "admin1@lshworld.com"  # Replace with your email
sender_password = "dpvqmxwsrxvxmbvr"  # Replace with your password or app password
receiver_email = ['cs4@lshworld.com']

logging.basicConfig(filename=r'C:\Users\USER\ImportOracle\pythonProject1\Log\Simplr_WS5_Import.log',level=logging.INFO,
                      # Set log level to DEBUG
                    format='%(asctime)s - %(levelname)s - %(message)s')
console_handler = logging.StreamHandler()  # Create a console handler
console_handler.setLevel(logging.DEBUG)  # Set console handler level to DEBUG
console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logging.getLogger('').addHandler(console_handler)

exceldata = r"C:\Users\USER\ImportOracle\pythonProject1\dist\data5.xlsx"

def GetLastestCustomer(exceldata):
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQFvubpNHb1TQEPliUeeuyqWx30SLFagXt8CTDt1L4y4O_PLTSKiqQulEdnbNG-GdWKteAd7ueLB9f4/pub?output=xlsx"
    response = requests.get(url)
    with open(exceldata, "wb") as file:
        file.write(response.content)

    data = pd.read_excel(exceldata, sheet_name=None)
    with open(exceldata, "wb") as file:
        file.write(response.content)




def Output(sheet,combine_df):
    # Download the Excel file from Google Sheets
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQFvubpNHb1TQEPliUeeuyqWx30SLFagXt8CTDt1L4y4O_PLTSKiqQulEdnbNG-GdWKteAd7ueLB9f4/pub?output=xlsx"
    response = requests.get(url)


    with open(exceldata, "wb") as file:
        file.write(response.content)


    # Load the Excel file and select the "Angus Steak House" sheet
    data = pd.read_excel(exceldata, sheet_name=sheet)

    # Promote headers
    #data.columns = data.iloc[0]
    # = data[1:]

    # Change column types
    data["Delivery Date 送货日期"] = pd.to_datetime(data["Delivery Date 送货日期"], errors="coerce",format="%d-%m-%Y")
    data["Submission ID"] = data["Submission ID"].astype(str)


    # Split 'My Products: Products' column by delimiter ")"
    data["My Products: Products"] = data["My Products: Products"].str.split(")")
    data = data.explode("My Products: Products")

    # Further split columns and clean data
    try:
        data[["Outlet 地址.1", "Outlet 地址.2" ,"Outlet 地址.3"]] = data["Outlet 地址"].str.split("-", n=2,expand=True)
    except:
        #no data
        return

    data["My Products: Products"] = data["My Products: Products"].str.strip()
    data[["My Products: Products.1", "My Products: Products.2"]] = data["My Products: Products"].str.split("-", n=1,
                                                                                                           expand=True)
    data[["My Products: Products.2.1", "My Products: Products.2.2"]] = data["My Products: Products.2"].str.split("(",
                                                                                                                 n=1,
                                                                                                                 expand=True)
    data[["My Products: Products.2.2.1", "My Products: Products.2.2.2", "My Products: Products.2.2.3"]] = data[
        "My Products: Products.2.2"].str.split(", ", expand=True)

    # Clean and rename columns
    data["My Products: Products.2.2.1"] = data["My Products: Products.2.2.1"].str.replace("Amount: ", "").str.replace(
        " SGD", "").astype(float)
    data["My Products: Products.2.2.2"] = data["My Products: Products.2.2.2"].str.replace("Quantity:", "").astype(float)
    data["My Products: Products.2.2.3"] = data["My Products: Products.2.2.3"].str.replace(": ", "")
    data = data.rename(columns={
        "Delivery Date 送货日期": "delivery_date_required",
        "Outlet 地址.1": "bill_to",
        "Outlet 地址.2": "business_outlet",
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
    data["quantity_supplier"] = data["quantity_required"]
    data["delivery_date_supplier"] = data["delivery_date_required"]






    data["buyer_code"] = ""
    data["order_date"] = pd.Timestamp.now()
    data["purchase_order_date"] = data["order_date"]
    data["specific_request"] = data["remark"]

    data["Supplier"] = "Lim Siang Huat Pte Ltd"

    #Change Date Format
    data['delivery_date_required'] = pd.to_datetime(data['delivery_date_required'],
                                                    errors='coerce')  # Convert to datetime

    #data['delivery_date_required'] = pd.to_datetime(data['delivery_date_required']).dt.date
    data['delivery_date_required'] = data['delivery_date_required'].dt.strftime('%d/%m/%Y')

    data['delivery_date_supplier'] = pd.to_datetime(data['delivery_date_supplier'],
                                                    errors='coerce')
    #data['delivery_date_supplier'] = pd.to_datetime(data['delivery_date_supplier']).dt.date
    data['delivery_date_supplier'] = data['delivery_date_supplier'].dt.strftime('%d/%m/%Y')

    data['order_date'] = pd.to_datetime(data['order_date'],
                                                    errors='coerce')
    #data['order_date'] = pd.to_datetime(data['order_date']).dt.date
    data['order_date'] = data['order_date'].dt.strftime('%d/%m/%Y')

    data['purchase_order_date'] = pd.to_datetime(data['purchase_order_date'],
                                        errors='coerce')
    #data['purchase_order_date'] = pd.to_datetime(data['purchase_order_date']).dt.date
    data['purchase_order_date'] = data['purchase_order_date'] .dt.strftime('%d/%m/%Y')

    #change Data Type
    try:
        data["business_outlet"] = data["business_outlet"].astype(int)
    except:
        print("sdf")
    data["bill_to"] = data["bill_to"].astype(int)


    #filter out NA qty
    data = data[data['quantity_required'].notna()]

    # Striping the Data for supplier item code
    data['supplier_item_code'] = data['supplier_item_code'].replace(" ", "", regex=True)

    data["po_no"] = data.apply(
        lambda row: row["po_no"] + "-F" if row["supplier_item_code"].startswith(
            ("FR", "ZF", "FSI", "ZKF", "FF","RMVE" ,"RMFR", "CH","UN","JOFRBT")) else row["po_no"] + "-D",
        axis=1)
    data["type"] = data["po_no"].apply(lambda x: 1019 if x.endswith("-F") else 1016)



    #Remove Old PO

    if combine_df is not None and not combine_df.empty:
        data = data.merge(combine_df, on='po_no', how='left')
        data = data[data['Custom'].isnull()]
        data = data.drop(columns=['Custom'])




    # Reorder columns
    column_order = [
        "po_no", "buyer_code", "business_outlet", "Supplier", "order_date", "supplier_item_code", "item_name", "uom",
        "quantity_required", "quantity_supplier", "weight", "unit_price", "amount_required", "amount_supplier",
        "delivery_date_required", "delivery_date_supplier", "specific_request", "purchase_order_date", "remark",
        "bill_to", "type"]

    data = data.reindex(columns=column_order, fill_value="")

    data.sort_values(by='po_no', inplace=True)


    logging.info(f"{sheet}  done")
    return data


import os


def combine_text_files(input_folder, output_file):
    """
    Combines all text files within a folder into a single output file.

    Args:
        input_folder: The path to the folder containing the text files.
        output_file: The path to the output file where the combined text will be written.
    """

    try:
        with open(output_file, 'w', encoding='utf-8') as outfile:  # Open in write mode ('w') and specify UTF-8 encoding
            for filename in os.listdir(input_folder):
                if filename.endswith(".txt"):  # Process only .txt files (you can customize this)
                    filepath = os.path.join(input_folder, filename)
                    try:
                        with open(filepath, 'r', encoding='utf-8') as infile:  # Open input file with UTF-8 encoding
                            for line in infile:
                                outfile.write(line)
                            outfile.write('\n')  # Add a newline between files (optional)
                    except UnicodeDecodeError:
                        logging.error(f"Skipping file {filename} due to encoding error.  Consider specifying correct encoding.")
                    except Exception as e:
                        logging.error(f"Error reading file {filename}: {e}")

    except FileNotFoundError:
        logging.error(f"Input folder '{input_folder}' not found.")
    except Exception as e:
        logging.error(f"An error occurred: {e}")


# More robust version that handles potential encoding issues and provides more informative error messages:


def combine_text_files_robust(input_folder, output_file):
    """Combines text files, handling encoding errors and providing more robust error messages."""

    try:
        with open(output_file, 'w', encoding='utf-8') as outfile:
            for filename in os.listdir(input_folder):
                if filename.endswith(".txt"):
                    filepath = os.path.join(input_folder, filename)
                    combined_successfully = False  # Flag to track successful combination
                    encodings_to_try = ['utf-8', 'latin-1', 'cp1252']  # Common encodings. Add more if needed.

                    for encoding in encodings_to_try:
                        try:
                            with open(filepath, 'r', encoding=encoding) as infile:
                                for line in infile:
                                    outfile.write(line)
                                outfile.write('\n')  # Add newline between files
                                combined_successfully = True
                                #logging.info(f"File {filename} combined successfully using {encoding} encoding.")
                                break  # Exit the encoding loop if successful
                        except UnicodeDecodeError:
                            logging.error(f"Decoding error with {encoding} for file {filename}. Trying another encoding...")
                        except Exception as e:
                            logging.error(f"Error reading file {filename} with {encoding} encoding: {e}")

                    if not combined_successfully:
                        logging.error(f"Failed to combine file {filename} after trying multiple encodings.")

    except FileNotFoundError:
        logging.info(f"Input folder '{input_folder}' not found.")
    except Exception as e:
        logging.error(f"An error occurred: {e}")

def movePOZg(path,archivepath):

    if os.path.exists(path):
        try:
            os.rename(path, archivepath)
            os.remove(path)
        except:
            os.remove(path)


def Output_WS(file_path):
    """Transforms data from a text file, extracting purchase order numbers.

    Args:
        file_path: The path to the text file.

    Returns:
        A pandas DataFrame containing unique purchase order numbers with a "Custom" column.
        Returns an empty DataFrame if file not found or error occurs.
    """
    try:
        if not os.path.exists(file_path):
            logging.info(f"Error: File not found at {file_path}")
            return pd.DataFrame()  # Return empty DataFrame

        with open(file_path, 'r', encoding='cp1252') as f:  # Specify encoding
            lines = f.readlines()

        data = []
        for line in lines:
            if "OrderNo :" in line:
                parts = line.split("OrderNo :")
                if len(parts) > 1:  #check if split was successful
                    order_info = parts[1].strip().split()
                    if order_info:  #check if order_info is not empty
                        po_no = order_info[0]  #Take the first element as PO number
                        data.append({"po_no": po_no})

        df = pd.DataFrame(data)

        # Remove duplicates
        df = df.drop_duplicates(subset='po_no')

        # Add "Custom" column
        df['Custom'] = 'done'
        logging.info("OutputWS done")

        return df

    except Exception as e:
        print(f"An error occurred: {e}")
        return pd.DataFrame()  # Return empty DataFrame on error

def send_email_with_attachment(sender_email, sender_password, receiver_email, subject, body):
    """Sends an email with an attachment."""

    # Create message object
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject

    # Attach body to email
    message.attach(MIMEText(body, 'plain'))  # Use 'html' for HTML content

    '''# Attach file
    if attachment_path:  # Only attach if a path is provided
        try:
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)

                # Add header with filename
                filename = os.path.basename(attachment_path)
                part.add_header("Content-Disposition", f"attachment; filename= {filename}")

                message.attach(part)
        except FileNotFoundError:
            logging.error(f"Error: Attachment file not found: {attachment_path}")
            return False'''  # Indicate failure

    try:
        # Connect to SMTP server
        with smtplib.SMTP("smtp.office365.com", 587) as server: # Or smtp.office365.com
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, receiver_email, message.as_string())
        logging.info(f"Email sent successfully to {receiver_email}!")
        return True # Indicate success
    except Exception as e:
        logging.error(f"Error sending email {receiver_email}: {e}")
        return False # Indicate failure

def Export_Query(file_path,receiver_email):

        # run SimplrIMport
        os.startfile(r"C:\Simplr\WhatsAPP_simplr\SimlprLSHImport.exe")

        for _ in tqdm(range(30), desc="Importing", unit="s"):
            time.sleep(1)

        try:
            # read by default 1st sheet of an excel file, IF NO ERROR, MEANS GOOD ELSE WILL NEED TO CHECK THE ERRORR CODE
            dataframe1 = pd.read_excel(file_path)
            AllerrorFiles = os.listdir(r"C:\Simplr\WhatsAPP_simplr\Log")

            if dataframe1.empty:
                movePOZg(path, archive)
                logging.info(f"{file_path} no data")


            elif any("ErrorLogCS" in filename for filename in AllerrorFiles):

                now = datetime.now()
                current_date = now.strftime("%d%m%Y")

                logfile = "ErrorLogCS" + current_date + ".txt"

                OracleErrorLogFile = os.path.join(r"C:\Simplr\WhatsAPP_simplr\Log", logfile)

                with open(OracleErrorLogFile, 'r') as file:
                    content = file.read()
                    logging.error(content)
                    for emails in receiver_email:
                        send_email_with_attachment(sender_email, sender_password, emails, "WhatsApp5", content)

                movePOZg(path, archive)
                os.remove(OracleErrorLogFile)

        except:
            logging.info(f" {path} Import Successful")
            now = datetime.now()
            current_date = now.strftime("%d%m") + "0" + now.strftime("%Y")
            current_time = now.strftime("%H_%M")

            OracleLogFile = os.path.join(r"C:\Simplr\WhatsAPP_simplr\Log", f"Oracle{current_date}.txt")
            OracleLogFileNew = r"C:\Feasibility\WhatsApp Order\Output WS\Oracle" + current_date + "_" + current_time + ".txt"
            os.rename(OracleLogFile, OracleLogFileNew)
            logging.info(f"Done moving log file to: {OracleLogFileNew}")






if __name__ == '__main__':

    GetLastestCustomer(exceldata)
    form = []
    all_sheets = pd.read_excel(exceldata, sheet_name=None)

    # Access individual sheets
    for sheet_name, dataframe in all_sheets.items():
        form.append(sheet_name)

    combine_text_files_robust(r"C:\Feasibility\WhatsApp Order\Output WS", r"C:\Users\USER\ImportOracle\pythonProject1\dist\combined_text_robust.txt")

    outputws = Output_WS(r"C:\Users\USER\ImportOracle\pythonProject1\dist\combined_text_robust.txt")

    all_data = []
    for sheet in form:
        if sheet == 'Create Restaurants ':
            print('sdf')
        data = Output(sheet, outputws)
        all_data.append(data)  # Append the DataFrame to the list

    # After the loop, concatenate the list of DataFrames
    all_data = pd.concat(all_data, ignore_index=True)

    if all_data.empty:
        sys.exit("No data available. Exiting...")

    # Save the transformed data to a new Excel file
    pozg_file = r"C:\Simplr\WhatsAPP_simplr\Import\PO_ZG.xlsx"

    all_data.to_excel(pozg_file, index=False)
    Export_Query(pozg_file,receiver_email)
    print("All Done Export")
