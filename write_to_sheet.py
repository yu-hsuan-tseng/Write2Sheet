'''
    Objective: Automatically read gmail from designated account and write links to google sheet
    Developer: Yu-Hsuan Tseng
'''

import pandas as pd 
import numpy as np 
import requests
import time
from datetime import date,timedelta
import os
import re
import zipfile
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import imaplib,email,os
from email.header import decode_header
from unicodedata import normalize
from email.header import make_header
from imap_tools import MailBox
import webbrowser
import schedule
from os import path
import gspread
from oauth2client.service_account import ServiceAccountCredentials 
import datetime



class Run:

    def __init__(self):
        # line notify token 
        self.token = "gk4ksj4GjKBIHVq6LNZGcont8CZgpHwig36iiGFmAY7"

        self.sd = date.today() - timedelta(days=1)
        self.sd= self.sd.strftime('%Y%m%d')
        self.server = "smtp.gmail.com"
        self.port = 587 
        self.isGMAIL = True
        self.DIRECTORY_OF_IMAGES = "./data"
        self.NAME_OF_DESTINATION_ARCHIVE = "data" 

        self.username = "seanforlineec@gmail.com"
        self.password = "lineshoppingtw"
        self.sender_ = self.username 
        self.isGMAIL = True

        self.auth_json_path = './drive_key.json'
        self.gss_scopes = ['https://spreadsheets.google.com/feeds']


    def auth_gss_client(self,path_, scopes):
        credentials = ServiceAccountCredentials.from_json_keyfile_name(path_, scopes)
        return gspread.authorize(credentials)


    def get_confirm_token(self,response):
        for key, value in response.cookies.items():
            if key.startswith('download_warning'):
                return value

        return None

    def save_response_content(self,response, destination):
        CHUNK_SIZE = 32768

        with open(destination, "wb") as f:
            for chunk in response.iter_content(CHUNK_SIZE):
                if chunk: # filter out keep-alive new chunks
                    f.write(chunk)
    def download_file_from_google_drive(self,id, destination):
        URL = "https://docs.google.com/uc?export=download"

        session = requests.Session()

        response = session.get(URL, params = { 'id' : id }, stream = True)
        token = self.get_confirm_token(response)

        if token:
            params = { 'id' : id, 'confirm' : token }
            response = session.get(URL, params = params, stream = True)

        self.save_response_content(response, destination) 


    def lineNotifyMessage(self,token,msg):
    
        headers = {"Authorization":"Bearer "+token,
                "Content-Type":"application/x-www-form-urlencoded"}
        
        payload = {'message':msg}
        r = requests.post("https://notify-api.line.me/api/notify", headers = headers, params = payload)
        
        return r.status_code    

    '''

    # From http://stackoverflow.com/questions/14568647/create-zip-in-python
    def zip(self,src, dst):
        zf = zipfile.ZipFile("%s.zip" % (dst), "w")
        src = os.path.abspath(src)
        for d, s, f in os.walk(src):
            for n in f:
                if re.match(r"^.*[.](csv|xlsx)$", n):
                    abs_name = os.path.abspath(os.path.join(d,n))
                    arc_name = abs_name[len(src) + 1:]
                    zf.write(abs_name, arc_name)
        zf.close()

    '''

    def to_sheet(self,urls):
        date_ = date.today()
        date_=date_.strftime('%Y-%m-%d')
        gss_client = self.auth_gss_client(self.auth_json_path, self.gss_scopes)
        spreadsheet_key_path_all = '1arW4FZpVa1E69FbGUI_mxCC2AgRczUVXBg0gD8fszlg'
        sheet = gss_client.open_by_key(spreadsheet_key_path_all).sheet1
        for u in urls:
            data = []
            data.append(date_)
            data.append(u)
            sheet.insert_row(data,2)

        
    def email_check(self):

        urls = []
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        imap.login(self.username,self.password)
        status,messages = imap.select("inbox")
        N = 100
        messages = int(messages[0])
        total = 0
        
        for i in range(messages,messages-N,-1):
            try:    
                res,msg = imap.fetch(str(i),"(RFC822)")
                for response in msg:
                    #print(response)
                    d = date.today() - timedelta(days=1)
                    d = d.strftime("%Y%m%d")
                    if isinstance(response, tuple):
                        
                        msg = email.message_from_bytes(response[1])
                        #msg = response[1]
                        subject = decode_header(msg["Subject"])[0][0]
                        if isinstance(subject,bytes):
                            subject = subject.decode()
                        if msg.is_multipart():
                            for part in msg.walk():
                                content_type = part.get_content_type()
                                content_disposition = str(part.get("Content_Disposition"))
                                try:
                                    body = part.get_payload(decode=True).decode()
                                except:
                                    pass
                                if content_type == "text/plain" and "attachment" not in content_disposition:
                                
                                    if "Your Unsampled report" in subject and d in subject and "for Jesse" not in subject:
                                    
                                        total+=1
                                        regex = r"(?i)\b((?:https?://docs.google.com|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'\".,<>?«»“”‘’]))"
                                        url = re.findall(regex,body)
                                        for u in url[1]:
                                            if u =="":
                                                pass
                                            else:
                                                urls.append(u)
                        else:
                            content_type = msg.get_content_type()
                            body = msg.get_payload(decode=True)
                            if content_type == "text/plain":
                                pass
                            if content_type == "text/html":
                                pass
            except Exception as e:
                print(e)
                pass
        
        try:
            imap.close()
            imap.logout()
        except Exception as e:
            print(e)

        if total<26:
            self.lineNotifyMessage(self.token, "Sean 哥 GA 今天"+str(total)+" 份已寄出")
            return urls,26-len(urls)
        else:
            print(total)
            self.lineNotifyMessage(self.token, "GA 寄送資料已完成")
            
            return urls,26-len(urls)






    def job(self,i):

        
        urls,num = self.email_check()
        try:
            self.to_sheet(urls)

        except: 
            print("Job has error ...")
            

