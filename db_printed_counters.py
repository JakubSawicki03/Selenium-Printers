from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
import mysql.connector as mysql
import time, os, shutil, csv, cred, openpyxl

HOST = cred.mySQL_host              # mySQL - server IP address/domain name
DATABASE = 'liczniki'               # mySQL - database name, if you want just to connect to MySQL server, leave it empty
USER = cred.mySQL_usr               # mySQL - remote user
PASSWORD = cred.mySQL_pass          # mySQL - remote user's password

# FUNCTIONS #
def getPrintersAddresses(): # get all printers addresses from mySQL database
    conn = mysql.connect(host=HOST, database=DATABASE, user=USER, password=PASSWORD)        # connect to MySQL server
    mycursor = conn.cursor()

    query = 'SELECT Adres_drukarki from drukarki'               # select ip addresses of each printer (needed for for_loop)
    mycursor.execute(query)                                     # execute query

    rows = mycursor.fetchall()
    for row in rows:
        row = str(row)
        printers_addresses.append(row[2:-3])                    # append ip address of each printer to list

    conn.close()                                                # close connection to database

def tempDir():              # make temp directory for xlsx files
    try:
        os.mkdir(f'{path}\\counters_excel')             # create temp_directory for excel files
    except:
        pass

def countersDir():          # make directory for counters
    try:
        os.mkdir(f'{path}\\counters')                   # create directory for counters
    except:
        pass

def currentMonth():         # get current month
    currentMonth = datetime.now().month
    currentYear = str(datetime.now().year)

    match(currentMonth):
        case 1: currentMonth = "styczen"
        case 2: currentMonth = "luty"
        case 3: currentMonth = "marzec"
        case 4: currentMonth = "kwiecien"
        case 5: currentMonth = "maj"
        case 6: currentMonth = "czerwiec"
        case 7: currentMonth = "lipiec"
        case 8: currentMonth = "sierpien"
        case 9: currentMonth = "wrzesien"
        case 10: currentMonth = "pazdziernik"
        case 11: currentMonth = "listopad"
        case 12: currentMonth = "grudzien"

    currentMonth = currentMonth + '_' + currentYear

    return currentMonth

def previousMonth():        # get previous month
    currentMonth = datetime.now().month
    currentYear = str(datetime.now().year)

    if currentMonth != 1:
        previousMonth = currentMonth - 1
        
    else:
        previousMonth = 12
        currentYear = datetime.now().year - 1
        currentYear = str(currentYear)

    match(previousMonth):
        case 1: previousMonth = "styczen"
        case 2: previousMonth = "luty"
        case 3: previousMonth = "marzec"
        case 4: previousMonth = "kwiecien"
        case 5: previousMonth = "maj"
        case 6: previousMonth = "czerwiec"
        case 7: previousMonth = "lipiec"
        case 8: previousMonth = "sierpien"
        case 9: previousMonth = "wrzesien"
        case 10: previousMonth = "pazdziernik"
        case 11: previousMonth = "listopad"
        case 12: previousMonth = "grudzien"

    previousMonth = previousMonth + '_' + currentYear

    return previousMonth

def isConnection():         # check printer's availability
    ping = os.system(f"ping {printer_ip[7:21]} -n 1")      # check availability of printer by sending ping
    if ping == 0:                                                  
        whichMenu()
    else:
        with pd.ExcelWriter(f'{path}\\counters_excel\\{printer_ip[7:21]}.xlsx') as writer:
            df_ip = pd.DataFrame(data={'Ip_drukarki':[printer_ip[7:21]]})            # dataFrame - IP ADDRESS
            df_printer_number = pd.DataFrame(data={'Nr_drukarki':[(printer_ip[7:21])[-2:]]})
            df_black_counter = pd.DataFrame(data={'Czarny':[int(0)]})        # dataFrame - black counter
            df_color_counter = pd.DataFrame(data={'Kolor':[int(0)]})         # dataFrame - color counter
            df_scan_counter = pd.DataFrame(data={'Skany':[int(0)]})          # final scan counter dataFrame

            combined_df = pd.concat([df_printer_number, df_ip, df_black_counter, df_color_counter, df_scan_counter], ignore_index=False, axis=1)
            combined_df.to_excel(writer, sheet_name=f'{printer_ip[7:21]}', index=False)     # write excel file from dataFrame (each printer has separated xlsx file)

def whichMenu():            # check which menu has the appropriate printer
    if printer_ip[-3:] == "xml":
        newMenu()
    elif printer_ip[-3:] == "tml":
        oldMenu()
    elif printer_ip[-9:] == 'observium':
        observiumPrinter()

def oldMenu():              # get counter from the old printer's menu
    with pd.ExcelWriter(f'{path}\\counters_excel\\{printer_ip[7:21]}.xlsx') as writer:
        driver.get(f"{printer_ip}")

        printerIpTable.append(printer_ip[7:21]); time.sleep(1)      # append ip_address to the list
        
        driver.find_element(By.ID, 'public').click()         # Select to log in as public user
        driver.find_element(By.XPATH, '//input[@value="Zaloguj się"]').click(); time.sleep(1)      # log in 
        driver.find_element(By.XPATH, '//a[contains(text(),"Informacje o interfejsie")]').click(); time.sleep(1)       # go to the IP_ADDRESS page

        ip_div = driver.find_element(By.XPATH, '//dd[contains(text(),"192.168.80.")]').text             # scrap dd that contains ip_address
        df_ip = pd.DataFrame(data={'Ip_drukarki':[ip_div]})        # dataFrame - IP ADDRESS 
        printer_number = (printer_ip[7:21])[-2:]          # get printer number
        df_printer_number = pd.DataFrame(data={'Nr_drukarki':[printer_number]})  # dataFrame - PRINTER NUMBER

        driver.find_element(By.XPATH, '//*[contains(text(),"Licznik")]').click()        # go to the counters page
        data_table = driver.find_element(By.XPATH, '//div[@class="page_main"]')         # scrap table from printer website
        data_table_html = data_table.get_attribute('outerHTML')                         # convert element into raw string
        df_table1 = pd.read_html(data_table_html)[1]                                    # dataFrame - 2nd table
        df_table2 = pd.read_html(data_table_html)[2]                                    # dataFrame - 3th table
        df_table3 = pd.read_html(data_table_html)[-1]                                   # dataFrame - the last one table

        if len(df_table1.columns) > 2:      # black and color toner
            black_counter = int(df_table1.iloc[0, 2]) + int(df_table2.iloc[0, 2]) + int(df_table3.iloc[0, 1])       # sum of black counters
            color_counter = int(df_table1.iloc[0, 1]) + int(df_table2.iloc[0, 1])                                   # sum of color counters
            df_black_counter = pd.DataFrame(data={'Czarny':[black_counter]})                                        # final black counter dataFrame
            df_color_counter = pd.DataFrame(data={'Kolor':[color_counter]})                                         # final color counter dataFrame
            df_scan_counter = pd.DataFrame(data={'Skany':[int(0)]})                                                   # final scan counter dataFrame
            combined_df = pd.concat([df_printer_number, df_ip, df_black_counter, df_color_counter, df_scan_counter], ignore_index=False, axis=1)   # concate above dataFrames into one

        elif (printer_ip[7:21])[-3:] == str(123):        # scans
            black_counter = int(df_table1.iloc[0, 1]) + int(df_table2.iloc[0, 1]) + int(df_table3.iloc[0, 1])       # sum of black counters
            scan_counter = int(df_table3.iloc[0,2])
            df_black_counter = pd.DataFrame(data={'Czarny':[black_counter]})                                        # final black counter dataFrame
            df_color_counter = pd.DataFrame(data={'Kolor':[int(0)]})                                                # final color counter dataFrame
            df_scan_counter = pd.DataFrame(data={'Skany':[scan_counter]})                                           # scan counter for printer x.x.x.123
            combined_df = pd.concat([df_printer_number, df_ip, df_black_counter, df_color_counter, df_scan_counter], ignore_index=False, axis=1)

        else:       # monochromatic (only black toner)
            black_counter = int(df_table1.iloc[0, 1]) + int(df_table2.iloc[0, 1]) + int(df_table3.iloc[0, 1])       # sum of black counters
            df_black_counter = pd.DataFrame(data={'Czarny':[black_counter]})                                        # final black counter dataFrame
            df_color_counter = pd.DataFrame(data={'Kolor':[int(0)]})                                                # final color counter dataFrame
            df_scan_counter = pd.DataFrame(data={'Skany':[int(0)]})                                                 # final scan counter dataFrame
            combined_df = pd.concat([df_printer_number, df_ip, df_black_counter, df_color_counter, df_scan_counter], ignore_index=False, axis=1)       # concatenate dataFrames into one

        combined_df.to_excel(writer, sheet_name=f'{printer_ip[7:21]}', index=False)     # write excel file from dataFrame (each printer has separated xlsx file)

def newMenu():              # get counter from the new printer's menu
    with pd.ExcelWriter(f'{path}\\counters_excel\\{printer_ip[7:21]}.xlsx') as writer:
        driver.get(f'{printer_ip}')

        printerIpTable.append(printer_ip[7:21]); time.sleep(1)               # append ip_address to the list

        driver.find_element(By.XPATH, '//input[@value="Html"]').click()         # view mode (flash)
        driver.find_element(By.XPATH, '//input[@value="Login"]').click(); time.sleep(2)      # log in  
        driver.find_element(By.XPATH, '//*[contains(text(),"Network Setting Information")]').click(); time.sleep(2)      # go to printer information (for IP address)
        
        ip_div = driver.find_element(By.ID, value='S_NET').get_attribute('outerHTML')        # scrap div that contains ip_address
        ip_table = pd.read_html(ip_div)[1]                  # dataFrame - DIV that contains ip adress
        ip_address = ip_table.iloc[2, 1]                    # select the second row (index 2) and the first column (index 1)
        df_ip = pd.DataFrame(data={'Ip_drukarki':[ip_address]})      # dataFrame - IP ADDRESS 
        printer_number = (printer_ip[7:21])[-2:]            # get printer number
        df_printer_number = pd.DataFrame(data={'Nr_drukarki':[printer_number]})              # dataFrame - PRINTER NUMBER
    
        driver.find_element(By.XPATH, '//*[contains(text(),"Device Information")]').click(); time.sleep(2)     #  click the counters menu
        driver.find_element(By.XPATH, '//*[contains(text(),"Meter Count")]').click()            #  go to the counters page

        data_table = driver.find_element(By.ID, value='S_COU')      # scrap table from printer website
        data_table_html = data_table.get_attribute('outerHTML')     # convert element into raw string
        df_table1 = pd.read_html(data_table_html)[1]                # dataFrame - 2nd tabbe
        df_table2 = pd.read_html(data_table_html)[3]                # dataFrame - 4th table
        df_table3 = pd.read_html(data_table_html)[-1]               # dataFrame - the last one table

        black_counter = int(df_table2.iloc[0, 2]) + int(df_table3.iloc[0, 2])   # sum of black counters
        color_counter = int(df_table1.iloc[0, 3]) + int(df_table2.iloc[0, 1]) + int(df_table2.iloc[1, 1]) + int(df_table3.iloc[0, 1]) + int(df_table3.iloc[0, 3])       # sum of color counters
        df_black_counter = pd.DataFrame(data={'Czarny':[black_counter]})        # dataFrame - black counter
        df_color_counter = pd.DataFrame(data={'Kolor':[color_counter]})         # dataFrame - color counter
        df_scan_counter = pd.DataFrame(data={'Skany':[int(0)]})                 # final scan counter dataFrame

        combined_df = pd.concat([df_printer_number, df_ip, df_black_counter, df_color_counter, df_scan_counter], ignore_index=False, axis=1)   # concatenate dataFrames into one
        combined_df.to_excel(writer, sheet_name=f'{printer_ip[7:21]}', index=False)        # write excel file from dataFrame (each printer has separated sheet)

def observiumPrinter():     # get counter from the observium website
    with pd.ExcelWriter(f'{path}\\counters_excel\\{printer_ip[7:21]}.xlsx') as writer:
        username = cred.usr_obs
        password = cred.pass_obs
        
        printerIpTable.append(printer_ip[7:21])

        df_ip = pd.DataFrame(data={'Ip_drukarki':[printer_ip[7:21]]})              # dataFrame - IP ADDRESS 
        df_printer_number = pd.DataFrame(data={'Nr_drukarki':[(printer_ip[7:21])[-2:]]})       # dataFrame - PRINTER NUMBER

        driver.get('https://192.168.0.34/'); time.sleep(0.2)
        driver.find_element(By.ID, value='details-button').click(); time.sleep(0.2)
        driver.find_element(By.ID, value='proceed-link').click(); time.sleep(0.2)

        driver.find_element(By.NAME, 'username').send_keys(username); time.sleep(0.2)   # credentials (login)
        driver.find_element(By.ID, 'password').send_keys(password); time.sleep(0.2)     # credentials (password)
        driver.find_element(By.ID, 'submit').click()                                    # submit credentials
        driver.get('https://192.168.0.34/device/device=14/tab=health/metric=counter/id=11/'); time.sleep(0.2)   # coutners for printer x.x.x.107

        black_counter = driver.find_element(By.XPATH, '//*[@id="main_container"]/div[1]/div[2]/div[6]/div/table/tbody/tr[1]/td[8]/strong/a').text   # get black counter from observium
        df_black_counter = pd.DataFrame(data={'Czarny':[black_counter]})        # dataFrame - black counter
        df_color_counter = pd.DataFrame(data={'Kolor':[int(0)]})                # dataFrame - color counter
        df_scan_counter = pd.DataFrame(data={'Skany':[int(0)]})                 # final scan counter dataFrame

        combined_df = pd.concat([df_printer_number, df_ip, df_black_counter, df_color_counter, df_scan_counter], ignore_index=False, axis=1)   # concatenate dataFrames into one
        combined_df.to_excel(writer, sheet_name=f'{printer_ip[7:21]}', index=False)        # write excel file from dataFrame (each printer has separated sheet)

def countersToMySql():      # open mySQL, create table, insert data, queries
    csvDataFrame = pd.read_csv(f'.\\{currentMonth}.csv', index_col=False, delimiter=',')
    df = pd.DataFrame(csvDataFrame)     # dataFrame from .csv file
    
    def convertToNull(value):           # convert value 0 to NULL
        if value == 0:
            return "NULL"
        else:
            return value

    df['Nr_drukarki'] = df['Nr_drukarki'].apply(convertToNull)  # check for 0 in df and change it to NULL
    df['Ip_drukarki'] = df['Ip_drukarki'].apply(convertToNull)  # check for 0 in df and change it to NULL
    df['Czarny'] = df['Czarny'].apply(convertToNull)            # check for 0 in df and change it to NULL
    df['Kolor'] = df['Kolor'].apply(convertToNull)              # check for 0 in df and change it to NULL
    df['Skany'] = df['Skany'].apply(convertToNull)              # check for 0 in df and change it to NULL

    conn = mysql.connect(host=HOST, database=DATABASE, user=USER, password=PASSWORD)        # connect to MySQL server
    mycursor = conn.cursor()
    
    # create a table
    createTable = f'CREATE TABLE {currentMonth} (Nr_drukarki INT(3) NOT NULL, Ip_drukarki VARCHAR(15) NOT NULL, Czarny INT(15) NULL, \
        Kolor INT(15) NULL, Skany INT(15) NULL, Stary_Czarny INT(15) NULL, Stary_Kolor INT(15) NULL, Stary_Skany INT(15) NULL);'           

    # drop a table
    dropTable = f'DROP table {currentMonth};'

    # insert counters from previous month to current month table
    insertIntoMonthBefore = f'UPDATE {currentMonth} \
        INNER JOIN {previousMonth} AS pM1 ON pM1.Nr_drukarki = {currentMonth}.Nr_drukarki \
        SET {currentMonth}.Stary_Czarny = pM1.Czarny, {currentMonth}.Stary_Kolor = pM1.Kolor, {currentMonth}.Stary_Skany = pM1.Skany \
        WHERE pM1.Nr_drukarki = {currentMonth}.Nr_drukarki;'       
    
    # query to select all printers with their latest counters and number of printed pages
    query = f'SELECT drukarki.Nr_drukarki, drukarki.Ip_drukarki, \
        {currentMonth}.Czarny, {currentMonth}.Stary_Czarny, ({currentMonth}.Czarny - {currentMonth}.Stary_Czarny) AS Roznica_Czarny, \
        {currentMonth}.Kolor, {currentMonth}.Stary_Kolor, ({currentMonth}.Kolor - {currentMonth}.Stary_Kolor) AS Roznica_Kolor, \
        {currentMonth}.Skany, {currentMonth}.Stary_Skany, ({currentMonth}.Skany - {currentMonth}.Stary_Skany) AS Roznica_Skany \
        FROM drukarki \
        LEFT JOIN {currentMonth} \
        ON drukarki.Nr_drukarki = {currentMonth}.Nr_drukarki;'
    
    
    try:    # create a table with counters from current month
        mycursor.execute(createTable)                                    # create a table
    except:
        mycursor.execute(dropTable)                                      # drop a table
        mycursor.execute(createTable)                                    # create a table

    for row in df.itertuples():                                     # insert latest counters to the current month table
        mycursor.execute(f'INSERT INTO {currentMonth} (Nr_drukarki, Ip_drukarki, Czarny, Kolor, Skany) \
            VALUES ({row.Nr_drukarki},"{row.Ip_drukarki}",{row.Czarny},{row.Kolor},{row.Skany})')
    conn.commit()                                                   # commit updated values

    mycursor.execute(insertIntoMonthBefore)                         # insert counters from previous month to the current month table
    conn.commit()                                                   # commit updated values
    
    mycursor.execute(query)                                         # select all printers with their latest counters and number of printed pages
    headings = ['Nr_drukarki','Ip_drukarki','Czarny','Stary_Czarny','Roznica_Czarny','Kolor','Stary_Kolor','Roznica_Kolor','Skany','Stary_Skany','Roznica_Skany'] # headings
    rows = mycursor.fetchall()                                      # read all rows from query

    fp = open(f'counters\\{currentMonth}.csv', 'w')                         # create .csv file and write to it data from query
    with open(f'counters\\{currentMonth}.csv', 'w', newline='') as fp:    
        myFile = csv.writer(fp)                                             # Create a CSV writer object
        myFile.writerow(headings)
        for row in rows:                                                # Write the data row by row into the CSV file
            myFile.writerow(row)                                                     

    conn.close()

def countersToExcel():      # convert csv to xlsx as final file
    read_file = pd.read_csv(f'.\\counters\\{currentMonth}.csv')
    read_file.to_excel(f'.\\counters\\{currentMonth}.xlsx', index=None, header=True)

    source_filename = f'.\\counters\\{currentMonth}.xlsx'
    destination_filename = 'Liczniki.xlsx'
    data_to_copy = pd.read_excel(source_filename)

    with pd.ExcelWriter(destination_filename, engine='openpyxl', mode='a') as writer:           # append data to destination file
        data_to_copy.to_excel(writer, sheet_name=currentMonth, index=False)

    
    workbook = load_workbook("Liczniki.xlsx")  # Load the workbook with .xlsm extension         # load the workbook

    # ADD CURRENT_MONTH TO TABLE FOR DATA VALIDATION
    sheet = workbook['listaArkusze']                                                            # Select the specific sheet by name

    row = 3
    while sheet[f"B{row}"].value is not None:                                                   # find the next available row in column B (starting from row 3)
        row += 1

    sheet[f"B{row}"] = currentMonth                                                             # set the value of the next available row in column B to currentMonth

    # ADD DATA VALIDATION
    sheet = workbook['liczniki_do_wyslania']                                                    # select the specific sheet by name

    formula_text = '=OFFSET(listaArkusze!$B$2,1,0, COUNTA(listaArkusze!$B$2:$B$100),1)'
    dv = DataValidation(type="list", formula1=formula_text, allow_blank=True)                   # set data validation for cell E3
    dv.add(sheet['E3'])
    sheet.add_data_validation(dv)

    # SAVE EXCEL FILE
    workbook.save("Liczniki.xlsx")                                                              # save the workbook with .xlsm extension  

    os.remove(f'.\\{currentMonth}.csv')                         # remove .csv file in main directory
    os.remove(f'.\\counters\\{currentMonth}.csv')              # remove .csv file in 'counters' directory

def mailMessage():          # send mail message to specified user with attachment
    import smtplib, ssl, email
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders

    smtp_server = cred.smtp_server
    smtp_port = 465                                             # SSL port
    smtp_username = cred.smtp_username                          # SMPT USER
    smtp_password = cred.smtp_password                          # SMPT PASSWORD
    sender_email = cred.smtp_username                           # SENDER
    recipient_email = cred.smtp_username                        # RECIEVER

    message = MIMEMultipart()                                   # Create a message
    message['From'] = sender_email                              # FROM
    message['To'] = recipient_email                             # TO
    message['Subject'] = 'Liczniki do wysłania'                 # SUBJECT

    attachment_file = f"{currentMonth}.xlsx"               # Attachment file name and path
    attachment_path = f"{path}\\counters\\{currentMonth}.xlsx"       # Attachment file name and path

    attachment = open(attachment_path, 'rb')                                                # open the file (rb - read binary)
    base = MIMEBase('application', 'octet-stream')                                          # open the file in binary mode
    base.set_payload(attachment.read())                                                     # read and set as the payload
    encoders.encode_base64(base)                                                            # encode base file
    base.add_header('Content-Disposition', f'attachment; filename={attachment_file}')   # give attachment name as varaible ATTACHMENT_FILE
    message.attach(base)                                                                    # attach the file

    context = ssl.create_default_context()                                  # Connect to the SMTP server with SSL
    server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=context)      # Connect to the SMTP server with SSL

    server.login(smtp_username, smtp_password)                              # Login to your email account
    server.sendmail(sender_email, recipient_email, message.as_string())     # Send the email
    server.quit()                                                           # Close the connection



# SCRIPT # 
driver = webdriver.Chrome()
path = os.getcwd()                  # get current path
printerIpTable = []                 # make a table to store all printer ip addresses
finalDf = []                        # make a list for sheets data frames (printer_number | printer_ip | black_counter | color_counter | scan_counter)
currentMonth = currentMonth()       # get current month
previousMonth = previousMonth()     # get previous month 
printers_addresses = []             # list to store printer's ip addresses
getPrintersAddresses()              # get all url addresses 
countersDir()                       # create foleder to store printers counters
tempDir()                           # create temp folder for temp excel files

for printer_ip in printers_addresses:
    printer_ip = printer_ip.strip()
    isConnection()

xlsxList = os.listdir(f'{path}\\counters_excel\\')      # check for all .xlsx files

for xlsx in xlsxList:                                   # for loop - create dataFrames from .xlsx files, then append them to the final version dataFrame
    if xlsx.endswith('.xlsx'):
        df = pd.read_excel(f'{path}\\counters_excel\\{xlsx}')
        finalDf.append(df)

printerIpTable.sort()                                   # sort the list to make it compatible with the dataframe
shutil.rmtree(f'{path}\\counters_excel')                # remove temp folder
result = pd.concat(finalDf, ignore_index=True)          # concatenate all .xlsx dataFrames into one main file
result.to_csv(f'{currentMonth}.csv', index=False)       # export final version of .xlsx to .csv 
driver.quit()

countersToMySql()                                       # create table and insert data into mySQL database
countersToExcel()                                       # convert .csv to .xlsx and add new sheet to the main file
mailMessage()                                           # send mail with counters.xlsx