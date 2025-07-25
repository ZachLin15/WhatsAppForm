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
    url= "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ6DOQzaBXUQ24qx2tOlE1sjK3ZEBdYxAoAbudQVNLl6GvVPqgza5QmMUMZhaU4vUYsb7rpuaJ3W4tN/pub?output=xlsx"
    response = requests.get(url)
    with open(exceldata, "wb") as file:
        file.write(response.content)

    data = pd.read_excel(exceldata, sheet_name=None)
    with open(exceldata, "wb") as file:
        file.write(response.content)




def Output(sheet,combine_df):
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ6DOQzaBXUQ24qx2tOlE1sjK3ZEBdYxAoAbudQVNLl6GvVPqgza5QmMUMZhaU4vUYsb7rpuaJ3W4tN/pub?output=xlsx"

    try:
        # Fetch data from Google Sheets
        response = requests.get(url)
        response.raise_for_status()

        # Read the specific sheet "Eric Ordering Form"
        with open(exceldata, "wb") as file:
            file.write(response.content)

        data = pd.read_excel(exceldata, sheet_name=sheet)


        # Promote headers (assuming first row is headers)
        if data.empty:
            return pd.DataFrame()

        # Convert Submission Date to datetime if it exists
        if 'Submission Date' in data.columns:
            data['Submission Date'] = pd.to_datetime(data['Submission Date'], errors='coerce')

        # Unpivot columns (melt the dataframe)
        df_melted = data.melt(var_name='Attribute', value_name='Value')

        # Split Attribute by '>>' - First split
        df_melted[['Attribute.1', 'temp1']] = df_melted['Attribute'].str.split('>>', n=1, expand=True)

        # Split temp1 by '>>' - Second split
        df_melted[['Attribute.2', 'Attribute.3']] = df_melted['temp1'].str.split('>>', n=1, expand=True)

        # Split Attribute.2 by '>>' again
        df_melted[['Attribute.2.1', 'Attribute.2.2']] = df_melted['Attribute.2'].str.split('>>', n=1, expand=True)

        # Merge Attribute.2.2 and Attribute.3
        df_melted['Attribute.3.1'] = df_melted['Attribute.2.2'].fillna('') + df_melted['Attribute.3'].fillna('')

        # Split item name by '-'
        df_melted[['Attribute.2.1', 'Attribute.2.2']] = df_melted['Attribute.2.1'].str.split('-', n=1, expand=True)

        # Convert Value to string and merge with Attribute.1
        df_melted['Value'] = df_melted['Value'].astype(str)
        df_melted['Merged'] = df_melted['Attribute.1'].fillna('') + df_melted['Value']

        # Split by various delimiters to extract order information
        df_melted[['Merged.1', 'Merged.2']] = df_melted['Merged'].str.split('Submission Date', n=1, expand=True)
        df_melted[['Merged.1.1', 'Merged.1.2']] = df_melted['Merged.1'].str.split('Outlet', n=1, expand=True)
        df_melted[['Merged.1.1.1', 'Merged.1.1.2']] = df_melted['Merged.1.1'].str.split('送货日期', n=1, expand=True)
        df_melted[['Merged.1.1.1.1', 'Merged.1.1.1.2']] = df_melted['Merged.1.1.1'].str.split('Submission ID', n=1,
                                                                                              expand=True)

        # Extract numeric part (quantity) using character transition logic
        df_melted['quantity'] = df_melted['Merged.1.1.1.1'].str.extract(r'(\d+)').astype('Int64')

        # Replace empty strings with null in Merged.1.2
        df_melted['Merged.1.2'] = df_melted['Merged.1.2'].replace('', np.nan)

        # Fill missing values up and down
        df_melted['Merged.1.1.1.2'] = df_melted['Merged.1.1.1.2'].fillna(method='ffill')
        df_melted['Merged.1.1.2'] = df_melted['Merged.1.1.2'].fillna(method='ffill')
        df_melted['Merged.1.2'] = df_melted['Merged.1.2'].fillna(method='ffill')
        df_melted['Merged.2'] = df_melted['Merged.2'].fillna(method='ffill')

        # Filter rows where Attribute.2.2 is not null
        df_filtered = df_melted[df_melted['Attribute.2.2'].notna()].copy()

        # Add index
        df_filtered.reset_index(drop=True, inplace=True)
        df_filtered['Index'] = df_filtered.index

        # Split item details by '/'
        df_filtered[['Attribute.2.1.1', 'Attribute.2.1.2', 'Attribute.2.1.3']] = df_filtered['Attribute.2.1'].str.split(
            '/', expand=True)

        # Split by '$'
        df_filtered[['Attribute.2.1.2.1', 'Attribute.2.1.2.2']] = df_filtered['Attribute.2.1.2'].str.split('$', n=1,
                                                                                                           expand=True)

        # Split by ') '
        df_filtered[['Attribute.2.1.2.1.1', 'Attribute.2.1.2.1.2']] = df_filtered['Attribute.2.1.2.1'].str.split(') ',
                                                                                                                 n=1,
                                                                                                                 expand=True)
        df_filtered[['Attribute.2.1.3.1', 'Attribute.2.1.3.2']] = df_filtered['Attribute.2.1.3'].str.split(') ', n=1,
                                                                                                           expand=True)

        # Filter rows where Attribute.2.1.2.1.1 is not null
        df_filtered = df_filtered[df_filtered['Attribute.2.1.2.1.1'].notna()].copy()

        # Split by '(' and remove first part
        df_filtered[['temp', 'Attribute.2.1.1.2']] = df_filtered['Attribute.2.1.1'].str.split('(', n=1, expand=True)

        # Convert price to numeric (Currency type)
        df_filtered['unit_price_base'] = pd.to_numeric(df_filtered['Attribute.2.1.1.2'], errors='coerce')

        # Process outlet information - split by '-'
        outlet_split = df_filtered['Merged.1.2'].str.split('-', expand=True)
        if len(outlet_split.columns) >= 2:
            df_filtered['bill_to'] = pd.to_numeric(outlet_split[0], errors='coerce').fillna(0).astype('Int64')
            df_filtered['ship_to'] = pd.to_numeric(outlet_split[1], errors='coerce').fillna(0).astype('Int64')
        else:
            df_filtered['bill_to'] = 0
            df_filtered['ship_to'] = 0

        # Merge UOM columns
        df_filtered['uomu'] = df_filtered['Attribute.2.1.2.1.1'].fillna('') + df_filtered['Attribute.2.1.3.1'].fillna(
            '')

        # Merge item code
        df_filtered['item_code'] = df_filtered['Attribute.2.1.2.1.2'].fillna('') + df_filtered[
            'Attribute.2.1.3.2'].fillna('')
        df_filtered['item_code'] = df_filtered['item_code'].str.replace(' ', '')

        # Rename columns
        df_filtered['po_no'] = df_filtered['Merged.1.1.1.2']
        df_filtered['delivery_date'] = pd.to_datetime(df_filtered['Merged.1.1.2'], errors='coerce').dt.date
        df_filtered['order_date'] = pd.to_datetime(df_filtered['Merged.2'], errors='coerce').dt.date

        # Add conditional UOM logic
        df_filtered['uom'] = df_filtered.apply(
            lambda row: row['uomu'] if 'EA' in str(row['Attribute.3.1']) else 'CT', axis=1
        )

        # Add category based on item code prefix
        df_filtered['Custom1'] = df_filtered['item_code'].apply(
            lambda x: 'F' if any(str(x).startswith(prefix) for prefix in ['FR', 'FSI', 'ZF', 'RMFR']) else 'D'
        )

        # Process unit price with conditional logic
        df_filtered['unit_price_ct'] = pd.to_numeric(df_filtered['Attribute.2.1.2.2'], errors='coerce')
        df_filtered['unit_price'] = df_filtered.apply(
            lambda row: row['unit_price_ct'] if 'CT' in str(row['Attribute.3.1']) and pd.notna(row['unit_price_ct'])
            else row['unit_price_base'] if pd.notna(row['unit_price_base'])
            else 0.0, axis=1
        )

        # Calculate amount required
        df_filtered['amount_required'] = df_filtered['quantity'] * df_filtered['unit_price']

        # Select and clean final columns
        final_columns = [
            'quantity', 'po_no', 'delivery_date', 'order_date',
            'bill_to', 'ship_to', 'item_code', 'uom', 'unit_price',
            'Custom1', 'amount_required', 'Attribute.3.1'
        ]

        # Rename Custom1 to category
        df_filtered = df_filtered.rename(columns={'Custom1': 'category'})

        # Select final columns
        df_final = df_filtered[final_columns].copy()

        # Clean data types
        df_final['quantity'] = pd.to_numeric(df_final['quantity'], errors='coerce').fillna(0).astype(int)
        df_final['unit_price'] = pd.to_numeric(df_final['unit_price'], errors='coerce').fillna(0.0)
        df_final['bill_to'] = pd.to_numeric(df_final['bill_to'], errors='coerce').fillna(0).astype(int)
        df_final['ship_to'] = pd.to_numeric(df_final['ship_to'], errors='coerce').fillna(0).astype(int)
        df_final['po_no'] = df_final['po_no'].astype(str)

        return df_final

    except Exception as e:
        print(f"Error processing data: {str(e)}")
        return pd.DataFrame()

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
        if sheet == "Eric Ordering Form":
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
