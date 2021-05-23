import requests
import os,sys
from bs4 import BeautifulSoup
import re
import time
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from tqdm import tqdm
import matplotlib.pyplot as plt
from selenium.webdriver.chrome.options import Options
import smtplib
from email.mime.text import MIMEText
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from tqdm import tqdm   
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd
from datetime import date,timedelta
import datetime
import os.path
from os import path
import schedule


line_url = "https://buy.line.me"
headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36'}
options = Options()
options.add_argument("user-data-dir={1}")
SCOPES = ['https://www.googleapis.com/auth/analytics.readonly']
KEY_FILE_LOCATION = './client_secrets 2.json' 
VIEW_ID = '160738988'
token = "gBchiEeYeZ96hok1J2L6ivWZiJXTpw6ay0h3Stx8WcZ"


def lineNotifyMessage(token,msg):
    
    headers = {"Authorization":"Bearer "+token,
               "Content-Type":"application/x-www-form-urlencoded"}
    
    payload = {'message':msg}
    r = requests.post("https://notify-api.line.me/api/notify", headers = headers, params = payload)
    
    return r.status_code        
    


def initialize_analyticsreporting():
  """Initializes an Analytics Reporting API V4 service object.
  Returns:
    An authorized Analytics Reporting API V4 service object.
  """
  credentials = Credentials.from_service_account_file(KEY_FILE_LOCATION, scopes=SCOPES)
  #service = build('sheets', 'v4', credentials=credentials)

  # Build the service object.
  analytics = build('analyticsreporting', 'v4', credentials=credentials)

  return analytics

# 待討論
def get_report(analytics):
    
    sd = date.today() - timedelta(days=2)
    sd = sd.strftime('%Y-%m-%d') 

    return analytics.reports().batchGet(
      body={
        'reportRequests':[
        {
          'viewId': VIEW_ID,
          'pageSize': 40000,
          'dateRanges': [{'startDate': sd, 'endDate': sd}],
         'metrics': [{'expression': 'ga:totalEvents'}],
        'dimensions': [{'name': 'ga:date'},{'name': 'ga:eventAction'}],
            'dimensionFilterClauses': [{"filters": [{'dimensionName': "ga:eventCategory",
                                          "operator": "IN_LIST",
                                          "expressions": ['transfer', 'Android-transfer', 'iOS-transfer', 'transfer_barcode']}]},
                                        {"filters": [{'dimensionName': "ga:eventLabel",
                                          "operator": "EXACT",
                                          "expressions": ['已登入-AUTO']}]},
    ],  
    }]}).execute()


def convert_to_dataframe(response):
    
  for report in response.get('reports', []):
    columnHeader = report.get('columnHeader', {})
    dimensionHeaders = columnHeader.get('dimensions', [])
    metricHeaders = [i.get('name',{}) for i in columnHeader.get('metricHeader', {}).get('metricHeaderEntries', [])]
    finalRows = []
    

    for row in report.get('data', {}).get('rows', []):
      dimensions = row.get('dimensions', [])
      metrics = row.get('metrics', [])[0].get('values', {})
      rowObject = {}

      for header, dimension in zip(dimensionHeaders, dimensions):
        rowObject[header] = dimension
        
        
      for metricHeader, metric in zip(metricHeaders, metrics):
        rowObject[metricHeader] = metric

      finalRows.append(rowObject)
      
      
  dataFrameFormat = pd.DataFrame(finalRows)    
  return dataFrameFormat

      
def print_response(response):
  """Parses and prints the Analytics Reporting API V4 response.
  Args:
    response: An Analytics Reporting API V4 response.
  """
  for report in response.get('reports', []):
      columnHeader = report.get('columnHeader', {})
      dimensionHeaders = columnHeader.get('dimensions', [])
      metricHeaders = columnHeader.get('metricHeader', {}).get('metricHeaderEntries', [])
      for row in report.get('data', {}).get('rows', []):
        dimensions = row.get('dimensions', [])
        dateRangeValues = row.get('metrics', [])
        for header, dimension in zip(dimensionHeaders, dimensions):
            print (header + ': ' + dimension)

        for i, values in enumerate(dateRangeValues):
            print('Date range: ' + str(i))
        for metricHeader, value in zip(metricHeaders, values.get('values')):
            print(metricHeader.get('name') + ': ' + value)


def ga():
    analytics = initialize_analyticsreporting()
    response = get_report(analytics)
    df = convert_to_dataframe(response)
    df = df.rename(columns={'ga:date':'Date'})
    df = df.rename(columns={'ga:totalEvents':'Total Events'})
    df = df.rename(columns={'ga:eventAction':'Event Action'})

    return df



def preprocessing(df):
    
    sd = date.today() - timedelta(days=2)
    sd = sd.strftime('%Y%m%d')
    data = {'Date':[],'Event Action':[],'Total Events':[]}
    all_sid = []
    for d in df['Event Action']:
        d = d.split(",")
        try:
            all_sid.append(d[1])
        except:
            all_sid.append("0")
    df['sid'] = all_sid
    
    
    usid = df.sid.unique()
    for u in usid:
        data['Date'].append(sd)
        data['Event Action'].append("shop_name,"+u)
        target = df.loc[df['sid'].isin([u])]
        data['Total Events'].append(sum(target['Total Events']))
    data = pd.DataFrame(data)
    data = data.sort_values(by='Total Events',ascending=False)
    return data


def send_csv(attachment):
    
    
    sd = date.today() - timedelta(days=2)
    sd = sd.strftime('%Y-%m-%d')
    emailfrom = "seanforlineec@gmail.com"
    emailto = "tw_line_shopping@linecorp.com"
    
    fileToSend = attachment
    username = "seanforlineec@gmail.com"
    password = " lineshoppingtw"

    msg = MIMEMultipart()
    msg["From"] = emailfrom
    msg["To"] = emailto
   
    msg["Subject"] = sd+"_GA_report"
    msg.preamble = sd+"_GA_report"
   
    
    ctype, encoding = mimetypes.guess_type(fileToSend)
    if ctype is None or encoding is not None:
        ctype = "application/octet-stream"

    maintype, subtype = ctype.split("/", 1)

    if maintype == "text":
        fp = open(fileToSend)
    # Note: we should handle calculating the charset
        attachment = MIMEText(fp.read(), _subtype=subtype)
        fp.close()
    else:
        fp = open(fileToSend, "rb")
        attachment = MIMEBase(maintype, subtype)
        attachment.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", "attachment", filename=fileToSend)
    msg.attach(attachment)

    server = smtplib.SMTP("smtp.gmail.com:587")
    server.starttls()
    server.login(username,password)
    server.sendmail(emailfrom, emailto, msg.as_string())
    server.quit()

    print("send csv successfully")
    
def send_mail(content):
    
    gmail_user = "seanforlineec@gmail.com"
    gmail_pwd = "lineshoppingtw"
    msg = MIMEText(content)
    msg['Subject'] = 'Tableau system report'
    msg['From'] = gmail_user
    msg['To'] = 'seanforlineec@gmail.com'
    
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(gmail_user, gmail_pwd)
    server.send_message(msg)
    server.quit()

    print("Email has successfully sent to designated receiver")
    



def main():
    
    d= date.today() - timedelta(days=2)
    date_ = d.strftime("%Y-%m-%d")
    file = "Tableau_Daily_Partner_Clickout TABLEAUTW "+date_+" "+date_+".csv"
       
    try:
        if path.exists(file):
            pass
        else:
            df = ga()
            df['Total Events'] = df['Total Events'].astype('int')
            df = df.sort_values(by='Total Events',ascending=False)
            df.index = np.arange(1,len(df)+1)
            
            # adding preprocessing function 
            df = preprocessing(df)
        
        
            df.to_csv(file,index=False,encoding="utf-8-sig")
            send_csv(file)
            send_mail("File - Tableau Daily Partner Clickout 1 Sent successfully ")
            msg = "順利寄出檔案 "+date_
            lineNotifyMessage(token,msg)
        
            
           
    except:
        print("Error in main function ...")
        pass
        
    


if __name__=="__main__":
    
    
    while True:    
        x = datetime.datetime.now()
        hr = x.strftime("%H")
        mini = x.strftime("%M")
        if hr == "08" and mini == "50":
            main()
        else:
            pass
    
    
    