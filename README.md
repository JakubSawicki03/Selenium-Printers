# Selenium-Printer
This Python script automates the process of gathering printer counters from various sources (web interfaces, observium), storing them in a MySQL database, and then exporting the data to an Excel file while also sending it as an email attachment.  
# How to Use: 
1. Make sure you have the required Python libraries installed (selenium, pandas, openpyxl, mysql-connector-python, time, os, shutil, csv).
2. Provide necessary credentials and adjust file paths and server configurations in the cred.py file.
3. Run the Python script (script_name.py).
4. The script will create an Excel file (Liczniki.xlsx) containing counters for different printers in separate sheets. Additionally, an email will be sent with this file as an attachment.

# Notes   
This script uses Selenium for web scraping printer interfaces, MySQL Connector for database operations, and Pandas for data manipulation and storage.<br>
Make sure to customize the script according to your specific environment, such as providing valid login credentials, URLs, and file paths.
