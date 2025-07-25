import os.path

import win32com.client
import time
import shutil
import subprocess
from tqdm import tqdm
from datetime import datetime
import logging
import pandas as pd
from pathlib import Path
import win32
import sys



import cx_Oracle as oracledb
import cx_Oracle
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

oracledb.init_oracle_client(lib_dir= r'C:\Users\USER\ImportOracle\pythonProject1\instantclient_23_5')

logging.basicConfig(filename=r'C:\Users\USER\ImportOracle\pythonProject1\Simplr_Import.log',level=logging.INFO,
                      # Set log level to DEBUG
                    format='%(asctime)s - %(levelname)s - %(message)s')
console_handler = logging.StreamHandler()  # Create a console handler
console_handler.setLevel(logging.DEBUG)  # Set console handler level to DEBUG
console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logging.getLogger('').addHandler(console_handler)

path = r"C:\Simplr\WhatsAPP_simplr\Import\PO_ZG.xlsx"
archive = r"C:\Simplr\WhatsAPP_simplr\Archive\PO_ZG.xlsx"







def wait_for_refresh(excel_app, timeout=15):
    """
    Waits for all queries in the active Excel workbook to refresh.

    Args:
        excel_app: The Excel application object.
        timeout: The maximum number of seconds to wait.

    Returns:
        True if the refresh completes within the timeout, False otherwise.
    """

    start_time = time.time()

    while True:
        # Check if all queries are refreshed

        print("waiting to Refresh")
        if excel_app.Application.CalculationState == -4105:  # xlDone
            logging.info("All queries refreshed successfully.")
            return True

        # Check if timeout has been reached
        if time.time() - start_time > timeout:
            logging.warning(f"Timeout reached after {timeout} seconds while waiting for refresh.")
            return False
def Export_query(file_path):

    folders = len(file_path)
    for files in os.listdir(file_path):
        if not files.startswith("~"):
            try:
                excel = win32com.client.Dispatch('Excel.Application')

                #excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
                excel.Visible = False  # Optional: Run Excel in the background

                excelfile = os.path.join(file_path,files)
                workbook = excel.Workbooks.Open(excelfile)

                logging.info(f"Exporting file: {files}")

                # Refresh all connections (including Power Queries)

                for connection in workbook.Connections:
                    if str(connection.Name).__contains__('Data')  :
                        print(f"Refreshing: {connection.Name}")
                        connection.Refresh()

                        for _ in tqdm(range(20), desc="Refreshing", unit="s"):
                            time.sleep(1)

                        break

                excel.Application.Run("ExportData")
                workbook.Save()
                workbook.Close()
                excel.Quit()

                try:
                    subprocess.run(["taskkill", "/f", "/im", "excel.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    logging.error(f"Error terminating Excel processes: {e}")

                #run SimplrIMport
                os.startfile(r"C:\Simplr\WhatsAPP_simplr\SimlprLSHImport.exe")

                for _ in tqdm(range(30), desc="Importing", unit="s"):
                    time.sleep(1)






                try:
                    # read by default 1st sheet of an excel file, IF NO ERROR, MEANS GOOD ELSE WILL NEED TO CHECK THE ERRORR CODE
                    dataframe1 = pd.read_excel(r"C:\Simplr\WhatsAPP_simplr\Import\PO_ZG.xlsx")
                    AllerrorFiles = os.listdir(r"C:\Simplr\WhatsAPP_simplr\Log")

                    if dataframe1.empty:
                        movePOZg(path, archive)
                        logging.info(f"{files} no data")


                    elif any("ErrorLogCS" in filename for filename in AllerrorFiles):

                        now = datetime.now()
                        current_date = now.strftime("%d%m%Y")

                        logfile = "ErrorLogCS"+current_date+".txt"

                        OracleErrorLogFile = os.path.join(r"C:\Simplr\WhatsAPP_simplr\Log", logfile)

                        with open(OracleErrorLogFile, 'r') as file:
                            content = file.read()
                            logging.error(content)
                        movePOZg(path, archive)
                        os.remove(OracleErrorLogFile)

                except:
                    logging.info(f" {files} Import Successful")
                    now = datetime.now()
                    current_date = now.strftime("%d%m") + "0" + now.strftime("%Y")
                    current_time = now.strftime("%H_%M")

                    OracleLogFile = os.path.join(r"C:\Simplr\WhatsAPP_simplr\Log", f"Oracle{current_date}.txt")
                    OracleLogFileNew = r"C:\Feasibility\WhatsApp Order\Output WS\Oracle" + current_date + "_" + current_time + ".txt"
                    os.rename(OracleLogFile, OracleLogFileNew)
                    logging.info(f"Done moving log file to: {OracleLogFileNew}")


            except Exception as e:
                logging.error(f"Error running macro: {e}")
                if e.name == 'CLSIDToClassMap':
                    mod_name = e.obj.__name__
                    mod_name_parts = mod_name.split('.')
                    if len(mod_name_parts) == 3:
                        # Deleting the problematic module cache folder
                        gen_path = Path(win32com.client.gencache.GetGeneratePath())
                        folder_name = mod_name_parts[2]

                        folder_path = gen_path.joinpath(folder_name)
                        shutil.rmtree(folder_path)
                        # Deleting the reference to the module to force a rebuilding by gencache
                        del sys.modules[mod_name]
                        continue
                else:
                    raise Exception("There was an error loading Excel.") from e
                    continue

    logging.info(f"Finished ALL exporting and importing data to Oracle.")

def oracle_check():
    #GET TODAY'S DATE
    now = datetime.now()
    current_date = now.strftime("%d%m%Y")
    #FOLDER FOR THE PO_ZG
    file_path = os.path.join(r"C:\Simplr\WhatsAPP_simplr\Archive", current_date, "PO_ZG.xlsx")
    unique_po_no_list=[]

    #READ THE PO_ZG TO GET THE LIST OF PO_NO
    try:
        df = pd.read_excel(file_path)
        unique_po_no_list = df['po_no'].unique().tolist()

    except Exception as e:
        logging.error("No File Found")

    #MAKE CONNECTION TO ORACLE DATABASE
    connection = oracledb.connect(user="apps", password="apps", dsn="192.168.200.179/erpp", encoding="UTF-8")

    #SQL COMMAND
    sql = "SELECT ORIG_SYS_DOCUMENT_REF from XXWMS_OE_HEADER_IFACE_ALL WHERE CUSTOMER_PO_NUMBER like '%' || :PONO || '%'"
    sql2 = """SELECT * from XXWMS_OE_LINES_IFACE_ALL WHERE  ORIG_SYS_DOCUMENT_REF= :d"""
    sql3 =  """SELECT OPT.* FROM OE_PROCESSING_MSGS OPM JOIN OE_PROCESSING_MSGS_TL OPT ON OPM.TRANSACTION_ID = OPT.TRANSACTION_ID    WHERE OPM.ORIGINAL_SYS_DOCUMENT_REF = :SO_NO """""

   #LOOP THRU THE  PO_NO lIST
    for PONO in unique_po_no_list:                                                                                                        #
        #Get All SO
        cursor = connection.cursor()
        # Execute the SELECT statements
        cursor.execute(sql, {"PONO": PONO})
        result = pd.DataFrame()
        result = cursor.fetchall()
        cleaned_result = []

        #REMOVED " , (, ) " FOR THE SONUMBER
        for r in result:
            r = str(r).replace(",","")
            r = str(r).replace("(","")
            r = str(r).replace(")","")
            r = str(r).replace("'", "")
            r = str(r).replace("'", "")
        cleaned_result.append(r)
        #orig_sys_doc_ref_list = [row[0] for row in cursor.fetchall()]
        cursor.close()


        for d in cleaned_result:
            cursor = connection.cursor()
            cursor.execute(sql2,{"d": d})
            result2 = cursor.fetchone()
            data2 = pd.DataFrame([result2])
            cursor.close()

            cursor = connection.cursor()
            cursor.execute(sql3, {"d": d})
            result3 = cursor.fetchone()
            data3 = pd.DataFrame([result3])
            cursor.close()












def movePOZg(path,archivepath):

    if os.path.exists(path):
        try:
            os.rename(path, archivepath)
            os.remove(path)
        except:
            os.remove(path)


def main():
  movePOZg(path,archive)

  # Path to your Excel file
  file_path = r"C:\Feasibility\WhatsApp Order\Queries"
  Export_query(file_path)





# Call the main function if this script is executed directly
if __name__ == "__main__":
  main()

  #oracle_check()


