
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter #PdfToText
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO 
import openpyxl #Excel
import re #Regex
import smtplib, ssl #email
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import sys
import pathlib
import ctypes
import datetime



FILE_NAME = "PLACEHOLDER"
EXCEL_FILE_NAME = "FILENAME"
EXCEL_FILE_PATH = pathlib.Path(pathlib.Path.cwd(), EXCEL_FILE_NAME)
EXCEL_APP_NUM_REGEX = r"[SWV][0-9]{9}" 
EXCEL_DATE_REGEX = r"[20][1-9]{2}"
EXCEL_SELECTED_SHEET = ''
EXCEL_FIRST_ROW = 2
EXCEL_LAST_ROW = 999
EXCEL_LAST_NAME = 2
EXCEL_FIRST_NAME = 3
EXCEL_DOB = 4
EXCEL_UCI = 5
EXCEL_NATIONALITY = 6
EXCEL_AGENT_NAME = 7
EXCEL_TYPE = 8
EXCEL_APP_NUM = 9
EXCEL_APP_SUBMISSION = 10
EXCEL_APP_REQUEST = 11
EXCEL_DECISION = 12
EXCEL_APPROVAL = 13
EXCEL_STATUS = 14
EXCEL_NOTE = 15
EXCEL_SHEETS = ["2020", "2019 (new)", "2018", "2017", "2016", "2015", "2014", "2013", "2012"]

CLIENT_APP_NUM = ''
CLIENT_AGENT_NAME = ''
CLIENT_FIRST_NAME = ''
CLIENT_LAST_NAME = ''
CLIENT_NATIONALITY = ''
CLIENT_TYPE = ''


GOOGLE_EMAIL = "ADDRESS"
GOOGLE_PW = "PW"
GOOGLE_AGENT_NAME = "James"
GOOGLE_DEPT1 = ["EMAIL1", "EMAIL2"]
GOOGLE_DEPT2 = ["EMAIL1", "EMAIL2"]
GOOGLE_DEPT3 = ["EMAIL1", "EMAIL2", "EMAIL3",]
GOOGLE_DEPT4 = ["EMAIL1", "EMAIL2", "EMAIL3", "EMAIL4"]
GOOGLE_DEPT5 = ["EMAIL1"]
GOOGLE_DEPT6 = ["EMAIL1", "EMAIL2", "EMAIL3"]
GOOGLE_DEPT7 = ["EMAIL1"]
GOOGLE_DEPT8 = ["EMAIL1"]

DEPT1_AGENT = ["AGENT1"]
DEPT2_AGENT = ["AGENT2"]
DEPT3_AGENT = ["AGENT3"]
DEPT4_AGENT = ["AGENT4"]
DEPT5_AGENT = ["AGENT5"]
DEPT6_AGENT = ["AGENT6"] 
DEPT7_AGENT = ["AGENT7"]
DEPT8_AGENT = ["AGENT8"]




def convert_pdf_to_txt(path):

    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 1
    caching = True
    pagenos=set()
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)
    fp.close()
    device.close()
    str = retstr.getvalue() # raw
    retstr.close()
    
    # Categorize docs
    if "5739" in str:
        client_status = "STATUS1"
    elif "5756" in str:
        client_status = "STATUS2"
    elif "5825" in str:
        client_status = "STATUS3"
    elif "5740" in str:
        client_status = "STATUS4"
    elif "5659" in str:
        client_status = "STATUS5"
    elif "5813" in str:
        client_status = "STATUS6"
    else:
        execute_alert('Warning', "The type of client's application is not detected or application was refused. Please process it manually", 0)
        exit()

    client_app_num = re.search(EXCEL_APP_NUM_REGEX, str).group()
    original_date = search_date(str)
    dateInFormat = convert_date(original_date)
    
    return {"num": client_app_num, "status": client_status, "date": dateInFormat, "path": path}


def execute_alert(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)


def modify_excel(client_info):
    sheets = openpyxl.load_workbook(EXCEL_FILE_PATH)
    selected_sheet = sheets[client_info["sheet"]]

    if "STATUS1" in client_info["status"]:
        client_info["note"] = client_info["note"] + "\n" + client_info["date"] + ": " + "STATUS1 letter"
    elif "STATUS2" in client_info["status"]:
        client_info["note"] = client_info["note"] + "\n" + client_info["date"] + ": " + "STATUS2 letter"
    elif "STATUS3" in client_info["status"]:
        client_info["note"] = client_info["note"] + "\n" + client_info["date"] + ": " + "STATUS3 letter"
    elif "STATUS4" in client_info["status"]:
        dttm = datetime.datetime.strptime(client_info["date"], "%Y-%m-%d")
        selected_sheet.cell(row=client_info["row"], column=EXCEL_DECISION).value = dttm
        client_info["note"] = client_info["note"] + "\n" + client_info["date"] + ": " + "STATUS4 letter"
    elif "STATUS5" in client_info["status"]:
        dttm = datetime.datetime.strptime(client_info["date"], "%Y-%m-%d")
        selected_sheet.cell(row=client_info["row"], column=EXCEL_APPROVAL).value = dttm
        client_info["note"] = client_info["note"] + "\n" + client_info["date"] + ": " + "STATUS5 letter"
    elif "STATUS6" in client_info["status"]:
        dttm = datetime.datetime.strptime(client_info["date"], "%Y-%m-%d")
        selected_sheet.cell(row=client_info["row"], column=EXCEL_APPROVAL).value = dttm
        client_info["note"] = client_info["note"] + "\n" + client_info["date"] + ": " + "STATUS6 letter"
    else:
        execute_alert('Not Found', "The status of client's application is not detected. Please process it manually", 0)
        exit()

    selected_sheet.cell(row=client_info["row"], column=EXCEL_NOTE).value = client_info["note"]
    sheets.save(EXCEL_FILE_PATH.resolve())


def modify_file_name(client_info): 

    last_name = client_info["last"].split(" ")[0]
    first_name = client_info["first"].split(" ")[0]
    
    if "STATUS1" in client_info["status"]:
        new_file_name = last_name + ", " + first_name + " - STATUS1 letter"
    elif "STATUS2" in client_info["status"]:
        new_file_name = last_name + ", " + first_name + " - STATUS2 letter"
    elif "STATUS3" in client_info["status"]:
        new_file_name = last_name + ", " + first_name + " - STATUS3 letter"
    elif "STATUS4" in client_info["status"]:
        new_file_name = last_name + ", " + first_name + " - STATUS4 letter"
    elif "STATUS5" in client_info["status"]:
        new_file_name = last_name + ", " + first_name + " - STATUS5 letter"
    elif "STATUS6" in client_info["status"]:
        new_file_name = last_name + ", " + first_name + " - STATUS6 letter"
    else:
        execute_alert('Not Found', "The status of client's application is not detected. Please process it manually", 0)
        exit()
    new_file_name = new_file_name.lower() + ".pdf"
    
    client_info["path"].rename(pathlib.Path(pathlib.Path.cwd(), new_file_name))
    client_info["file_name"] = new_file_name

    return client_info
    

def search_date(str):
    years = r'((?:19|20)\d\d)'
    pattern = r'(%%s) +(%%s), *%s' % years

    thirties = pattern % (
        "September|April|June|November",
        r'0?[1-9]|[12]\d|30')

    thirtyones = pattern % (
        "January|March|May|July|August|October|December",
        r'0?[1-9]|[12]\d|3[01]')

    fours = '(?:%s)' % '|'.join('%02d' % x for x in range(4, 100, 4))

    feb = r'(February) +(?:%s|%s)' % (
        r'(?:(0?[1-9]|1\d|2[0-8])), *%s' % years, # 1-28 any year
        r'(?:(29), *((?:(?:19|20)%s)|2000))' % fours)  # 29 leap years only

    EXCEL_DATE_REGEX = '|'.join('(?:%s)' % x for x in (thirties, thirtyones, feb))

    return re.search(EXCEL_DATE_REGEX, str).group()
    

def convert_date(dateInStr):
    
    dateInList = dateInStr.split(" ")
    
    if "January" in dateInStr:
        month = "01"
    elif "February" in dateInStr:
        month = "02"
    elif "March" in dateInStr:
        month = "03"
    elif "April" in dateInStr:
        month = "04"
    elif "May" in dateInStr:
        month = "05"
    elif "June" in dateInStr:
        month = "06"
    elif "July" in dateInStr:
        month = "07"
    elif "August" in dateInStr:
        month = "08"
    elif "September" in dateInStr:
        month = "09"
    elif "October" in dateInStr:
        month = "10"
    elif "November" in dateInStr:
        month = "11"
    elif "December" in dateInStr:
        month = "12"
    else:
        execute_alert('Date Not Found', "Cannot find a date in the letter. Please process it manually", 0)
        exit()
    return dateInList[2] + "-" + month + "-" + dateInList[1][:-1]



def read_excel(client_num):
    sheets = openpyxl.load_workbook(EXCEL_FILE_PATH)
    sheetNames = sheets.sheetnames
    for sheet_name in EXCEL_SHEETS:
        selectedSheet = sheets[sheet_name]
        for r in selectedSheet.iter_rows(min_row=EXCEL_FIRST_ROW, max_row=EXCEL_LAST_ROW, min_col=EXCEL_APP_NUM, max_col=EXCEL_APP_NUM):
            for cell in r:
                client_row = cell.row
                if cell.value == client_num:
                    EXCEL_SELECTED_SHEET = selectedSheet

                    client_last = EXCEL_SELECTED_SHEET.cell(row=client_row, column=EXCEL_LAST_NAME).value
                    client_first = EXCEL_SELECTED_SHEET.cell(row=client_row, column=EXCEL_FIRST_NAME).value
                    client_nationality = EXCEL_SELECTED_SHEET.cell(row=client_row, column=EXCEL_NATIONALITY).value
                    client_type = EXCEL_SELECTED_SHEET.cell(row=client_row, column=EXCEL_TYPE).value
                    client_agent = EXCEL_SELECTED_SHEET.cell(row=client_row, column=EXCEL_AGENT_NAME).value

                    client_note = EXCEL_SELECTED_SHEET.cell(row=client_row, column=EXCEL_NOTE).value
                    if not client_note:
                        client_note = ''
                        print("FINISHED read_excel")
                    return {"last": client_last, "first": client_first, "nationality": client_nationality, "type": client_type, "agent": client_agent, "row": client_row, "note": client_note, "sheet": sheet_name}

                elif cell.row >= EXCEL_LAST_ROW:
                    client_row = None
                    break
            else:
                continue
            break

        execute_alert('Not Found', "Client's application number is not found. Please process it manually", 0)
        exit()

 

def tuple_to_str(message):
    rst = ''.join(message)

    return rst

def create_email(last, first, nationality, agent, type, status):
    if "TYPE1" in type:
        type_for_msg = "TYPE1 - 1"
    elif "TYPE2" in type:
        type_for_msg = "TYPE1 - 2"
    elif "TYPE3" in type:
        type_for_msg = "TYPE3"
    elif "TYPE4" in type:
        type_for_msg = "TYPE4"
    elif "TYPE5" in type:
        type_for_msg = "TYPE5"
    elif "TYPE6" in type:
        type_for_msg = "TYPE6"
    elif "TYPE7" in type:
        type_for_msg = "TYPE7"
    else:
        execute_alert('Not Found', "The type of client's application is not detected. Please process it manually", 0)
        exit()

    if status == "STATUS1":
        result = ("Hello ",
                  agent,
                  ",\n\n",
                  "Please find the attached TYPE1 - 1 letter from ORGANIZATION NAME regarding ",
                  first, " ", last,
                  "'s ",
                  type_for_msg, " application. \n\n",
                  "We will notify you accordingly. \n\nBest regards,\n",
                  GOOGLE_AGENT_NAME)

    elif status == "STATUS2":
        result = ("Hello ",
                  agent,
                  ",\n\n",
                  "Please find the attached STATUS2 letter from ORGANIZATION NAME regarding ",
                   first, " ", last,
                   "'s ",
                   type_for_msg, " application. \n\n",
                   "We will notify you accordingly. \n\nBest regards,\n",
                   GOOGLE_AGENT_NAME)

    elif status == "STATUS3":
        result = ("Hello ",
                 agent,
                 ",\n\n",
                 "Please find the attached STATUS3 letter from ORGANIZATION NAME regarding ",
                 first, " ", last,
                 "'s ",
                 type_for_msg, " application. \n\n",
                 "We will notify you accordingly. \n\nBest regards,\n",
                 GOOGLE_AGENT_NAME)

    elif status == "STATUS4":
        result = ("Hello ",
                  agent,
                  ",\n\n",
                  "Please find the attached STATUS4 letter from ORGANIZATION NAME regarding ",
                  first, " ", last,
                  "'s ",
                  type_for_msg, " application. \n\n",
                  "We will notify you accordingly. \n\nBest regards,\n",
                  GOOGLE_AGENT_NAME)

    elif status == "STATUS5":
        result = ("Hello ",
                  agent,
                  ",\n\n",
                  "Please find the attached STATUS5 letter from ORGANIZATION NAME regarding ",
                  first, " ", last,
                  "'s ",
                  type_for_msg, " application. \n\n",
                  "We will notify you accordingly. \n\nBest regards,\n",
                  GOOGLE_AGENT_NAME)

    elif status == "STATUS6":
        result = ("Hello ",
                 agent,
                 ",\n\n",
                 "Please find the attached STATUS6 letter from ORGANIZATION NAME regarding ",
                 first, " ", last,
                 "'s ",
                 type_for_msg, " application. \n\n",
                 "We will notify you accordingly. \n\nBest regards,\n",
                 GOOGLE_AGENT_NAME)

    else:
        execute_alert('Not Found', "The status of client's application is not detected. Please process it manually", 0)
        exit()
   
    print("FINISHED create_email")
    result = tuple_to_str(result)
    return result    


def send_email(client_info):
    print("STARTED send_email")

    text = create_email(client_info["last"], client_info["first"], client_info["nationality"], client_info["agent"], client_info["type"], client_info["status"])

    port = 587  # For SSL
    smtp_server = "smtp.gmail.com" # Smtp
    sender_email = GOOGLE_EMAIL  # Enter your address
    receiver_email = "PLACEHOLDER"  # Enter receiver address
 
    
    if (client_info["agent"] in DEPT1_AGENT):
        receiver_email = GOOGLE_DEPT1
    elif (client_info["agent"] in DEPT2_AGENT):
        receiver_email = GOOGLE_DEPT2
    elif (client_info["agent"] in DEPT3_AGENT):
        receiver_email = GOOGLE_DEPT3
    elif (client_info["agent"] in DEPT4_AGENT):
        receiver_email = GOOGLE_DEPT4
    elif (client_info["agent"] in DEPT5_AGENT):
        receiver_email = GOOGLE_DEPT5
    elif (client_info["agent"] in DEPT6_AGENT):
        receiver_email = GOOGLE_DEPT6
    elif (client_info["agent"] in DEPT7_AGENT):
        receiver_email = GOOGLE_DEPT7
    elif (client_info["agent"] in DEPT8_AGENT):
        receiver_email = GOOGLE_DEPT8
    else:
        execute_alert('Not Found', "Unidentified agent. Email was not sent and Excel was not updated. Please check whether the agent's name and email address are added in a list", 0)
        exit()


    message = MIMEMultipart()
    message["From"] = GOOGLE_EMAIL
    message["To"] = "RECEIVER_EMAIL"
    message["Subject"] = "[" + client_info['status'].split(' ')[0] + "] " + client_info['last'] + ", " + client_info['first'] + " (" + client_info['type'] + ") - " + client_info['nationality']

    message.attach(MIMEText(text, "plain"))

    with open(client_info["path"].resolve(), "rb") as attachment:

        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
        "Content-Disposition",
        f"attachment; filename= {client_info['file_name']}",
    )
    message.attach(part)

    text = message.as_string()
    context = ssl.create_default_context()

    with smtplib.SMTP(smtp_server, port) as server:
        try:
            server = smtplib.SMTP(smtp_server, port)
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            server.login(GOOGLE_EMAIL, GOOGLE_PW)
            server.sendmail(GOOGLE_EMAIL, "RECEIVER_EMAIL", text) 
        except Exception as e:
            print(e)
            exit()
        finally:
            server.quit()

    return 0


if __name__ == "__main__":
    
    path ='%s' %sys.argv[1]  
    path = pathlib.Path(pathlib.Path.cwd(), path)
    client_info = convert_pdf_to_txt(path)
    
    more_client_info = read_excel(client_info["num"])
   
    client_info.update(more_client_info)

    client_info = modify_file_name(client_info)

    send_email(client_info)
    modify_excel(client_info)

    input('Press ENTER to exit')
    
