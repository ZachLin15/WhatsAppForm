import subprocess
import time
import logging

logging.basicConfig(filename=r'C:\Users\USER\ImportOracle\pythonProject1\Log\Main.log',level=logging.INFO,
                      # Set log level to DEBUG
                    format='%(asctime)s - %(levelname)s - %(message)s')

start_time = time.time()  # Record start time
"""try:
    print("Importing WhatsApp1")
    subprocess.run(['python', 'C:\\Users\\USER\\ImportOracle\\pythonProject1\\dist\\WhatApp1.py'])
except Exception as e:
        logging.error(f"WhatsApp1 {e}")

try:
    print("Importing WhatsApp2")
    subprocess.run(['python', 'C:\\Users\\USER\\ImportOracle\\pythonProject1\\dist\\WhatApp2.py'])
except Exception as e:
        logging.error(f"WhatsApp2 {e}")

try:
    print("Importing WhatsApp3")
    subprocess.run(['python', 'C:\\Users\\USER\\ImportOracle\\pythonProject1\\dist\\WhatApp3.py'])
except Exception as e:
        logging.error(f"WhatsApp3 {e}")

try:
    print("Importing WhatsApp5")
    subprocess.run(['python', 'C:\\Users\\USER\\ImportOracle\\pythonProject1\\dist\\WhatApp5.py'])
except Exception as e:
        logging.error(f"WhatsApp5 {e}")

try:
    print("Importing Yakun")
    subprocess.run(['python', 'C:\\Users\\USER\\ImportOracle\\pythonProject1\\dist\\Yakun.py'])
except Exception as e:
        logging.error(f"Yakun {e}")"""

try:
    print("Importing Eric Form")
    subprocess.run(['python', 'C:\\Users\\USER\\ImportOracle\\pythonProject1\\dist\\EricForm.py'])
except Exception as e:
        logging.error(f"Yakun {e}")


end_time = time.time()  # Record end time

elapsed_time = end_time - start_time
elapsed_time = elapsed_time / 60
print(f"Elapsed time: {elapsed_time:.4f} minute")
