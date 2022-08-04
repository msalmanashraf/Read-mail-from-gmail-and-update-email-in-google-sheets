#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Aug  3 23:06:11 2022

@author: salman
"""

from __future__ import print_function

import os.path
from apiclient import errors
import email
import base64
import pandas as pd
import json
from datetime import datetime
import datetime
import time
import sys

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from httplib2 import Http

# If modifying these scopes, delete the file token.json.
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/spreadsheets'
]

# here enter the id of your google sheet
SPREADSHEET_ID = '1n73oq9Os-Bz6XaZ75jm3QXadXN_Dq6SCq3QFpivbe4o'

current_sheet = 'Sheet1'
previous_sheet = 'Sheet1'

#specfic label of email you want to fetch from your email and update in google periodically
search_string = 'netflix'

#function to change current sheet and previous information because this progam add mail to monthly sheet.
def set_current_sheet(current_sheet, previous_sheet):
    
    #write json script
    dictionary ={
            'current' : current_sheet,
            'previous' : previous_sheet
            }
    
    #open json file in write mode and write json script
    with open("current_sheet_status.json", "w") as outfile:
        json.dump(dictionary, outfile)

#function to store current sheet information because this progam add mail to monthly sheet.
def get_current_sheet():
    
    #Opening JSON file in read mode
    with open('current_sheet_status.json', 'r') as openfile:
        
        # Reading from json file
        json_object = json.load(openfile)
        
    return json_object['current']

#function to store previous sheet information because this progam add mail to monthly sheet.
def get_previous_sheet():
    
    #Opening JSON file in read mode
    with open('current_sheet_status.json', 'r') as openfile:
        
        # Reading from json file
        json_object = json.load(openfile)
        
    return json_object['previous']

#google email and sheet service authentication function
def get_service(service_name):
    """
    Authenticate the google api client and return the service object 
    to make further calls
    PARAMS
        None
    RETURNS
        service api object from gmail for making calls
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    if (service_name == 'gmail'):
        service = build(service_name, 'v1', credentials=creds)
    elif (service_name == 'sheets'):
        service = build(service_name, 'v4', credentials=creds)

    return service

#search mail in unread with specific label and return list of msg Ids    
def search_message(service, user_id, search_string):
    
    """
    Search the inbox for emails using standard gmail search parameters
    and return a list of email IDs for each result
    PARAMS:
        service: the google api service object already instantiated
        user_id: user id for google api service ('me' works here if
        already authenticated)
        search_string: search operators you can use with Gmail
        (see https://support.google.com/mail/answer/7190?hl=en for a list)
    RETURNS:
        List containing email IDs of search query
    """

    try:
        # initiate the list for returning
        list_ids = []

        # get the id of all messages that are in the search string
        search_ids = service.users().messages().list(userId=user_id, labelIds = ['UNREAD'], maxResults=600, q=search_string).execute()
#        search_ids = service.users().messages().list(userId=user_id, q=search_string).execute()
        
        # if there were no results, print warning and return empty string
        try:
            ids = search_ids['messages']

        except KeyError:
            print("WARNING: the search queried returned 0 results")
            print("returning an empty string")
            return list_ids

        if len(ids)>1:
            for msg_id in ids:
                list_ids.append(msg_id['id'])
            return list_ids

        else:
            list_ids.append(ids[0]['id'])
            return list_ids
        
    except errors.HttpError as error:
        print("An error occured: %s") % error

#function define properties of sheet
def sheetProperties(title, **sheetProperties):
    defaultProperties = {
        'properties': {
            'title' : title,
            'index' : 0,
            'sheetType' : 'GRID',
            'hidden' : False
        }
    }
    defaultProperties['properties'].update(sheetProperties)
    return defaultProperties

#add new sheet in a spreadsheet 
def add_new_sheet(sheet, current_sheet, sheet_index):
    
    
    request_body = {
    'requests' :[
        {
           'addSheet' : sheetProperties(current_sheet)
            }
        ]
    }
    
    response = sheet.batchUpdate(
            spreadsheetId = SPREADSHEET_ID,
            body=request_body
    ).execute()
    
    values = (('S.No', 'Date', 'Time', 'Subject', 'Message'),
              )
    
    value_range_body = {
        'majorDimension': 'ROWS',
        'values': values
    }
    
    sheet.values().update(
        spreadsheetId = SPREADSHEET_ID,
        valueInputOption = 'USER_ENTERED',
        range= current_sheet + '!A' + str(sheet_index),   
        body=value_range_body
    ).execute()

#add new email to sheet    
def add_mail_to_sheet(sheet, current_sheet, sheet_index, mail_date, mail_time, subject, main_message):
    
    values = (
            (str(sheet_index-1), mail_date, mail_time, subject, main_message),
    )
    
    value_range_body = {
        'majorDimension': 'ROWS',
        'values': values
    }
    
    sheet.values().update(
        spreadsheetId = SPREADSHEET_ID,
        valueInputOption = 'USER_ENTERED',
        range= current_sheet + '!A' + str(sheet_index),   
        body=value_range_body
    ).execute()
    
    time.sleep(1)

# method get sheet index (number of entries) in present sheet
def get_sheet_index(sheet,current_sheet):
    
    RANGE_NAME = current_sheet +'!A1:A2000'
    result_input = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_NAME).execute()
    values_input = result_input.get('values', [])
    
    sheet_index =len(values_input) + 1
    
    return sheet_index

#convert epoch time to date and time
def get_date_and_time(epoch_time):
    
    #add offset due to my country time. I have offset 12 hours from standard time (12 * 3600000) 
    my_epoch_time = epoch_time + (43200000)
    
    # using the datetime.fromtimestamp() function  
    date_time = datetime.datetime.fromtimestamp( my_epoch_time/1000 )
    date = date_time.date()
    time = date_time.time()
    
    return date, time
 
#read mail from gmail using specific msg id
def get_message(service, user_id, msg_id):
    
    """
    Search the inbox for specific message by ID and return it back as a 
    clean string. String may contain Python escape characters for newline
    and return line. 
    
    PARAMS
        service: the google api service object already instantiated
        user_id: user id for google api service ('me' works here if
        already authenticated)
        msg_id: the unique id of the email you need
    RETURNS
        A string of encoded text containing the message body
    """
    try:
        # grab the message instance
        full_raw_message= service.users().messages().get(userId=user_id, id=msg_id, format="full", metadataHeaders=None).execute()
        
        # get subject of mail
        headers = full_raw_message["payload"]["headers"]
        subject = [i['value'] for i in headers if i["name"]=="Subject"]
        subjects = subject[0]
        
        # get main message of mail
        main_message = full_raw_message['snippet']
        
        # get desired data from raw format data here you can find desired data which you want extract from mail like(links, html formats and much more)
        '''raw_data=full_raw_message["payload"]["body"]['data']
        msg_str = base64.urlsafe_b64decode(raw_data)
        msg_str2 = msg_str.decode('UTF-8')'''

        
        #get date and time of mail
        email_date, email_time = get_date_and_time(int(full_raw_message['internalDate']))
        
        return email_date, email_time, subjects, main_message
        
    except Exception as error:
        print("An error occured: %s") % error    

#method to mark msg read(this method is used to mark msg read after adding mail to )   
def mark_unread(service, user_id, msg_id):
    service.users().messages().modify(userId=user_id, id=msg_id, body={
                'removeLabelIds': ['UNREAD']
                }).execute()


def main():

    #get gmail service using get_service function
    gmail_service = get_service('gmail')
    
    #get sheet service using get_service function
    sheet_service = get_service('sheets')
    sheet = sheet_service.spreadsheets()
    
    #get current sheet in which previous email are added
    current_sheet = get_current_sheet()
    
    #get previous_sheet in which previous month email were added
    previous_sheet = get_previous_sheet()
    
    #search indox using search message function(search function search messages in unread messages with specifc label.Specific label "search_string" is provided by user" ) 
    list_msg_Ids = search_message(gmail_service, 'me', search_string)
    
    #reverse the list of msg ids to get the most recent msg to be added in last 
    list_msg_Ids.reverse()

    #apply loop on msg ids to be put on google sheet
    for msg_Ids in list_msg_Ids:
        #read(fetch) email using msg id
        email_date, email_time, subject, main_message = get_message(gmail_service, 'me', msg_Ids)
        #convert email date to month-year format
        email_month_year = email_date.strftime("%m-%Y")
        #convert email date to day-month-year format to be put on sheet
        email_date_str = email_date.strftime("%d-%m-%Y")
        #convert email date to hour-minute-second format to be put on sheet
        email_time_str = email_time.strftime("%H:%M:%S")
        #now check if current sheet and current email have same month-year if not new sheet is created 
        if(email_month_year == current_sheet):
            #get on which number of row on current sheet new entry should put
            sheet_index = get_sheet_index(sheet, current_sheet)
            #add mail to existing sheet
            add_mail_to_sheet(sheet, current_sheet, sheet_index, email_date_str, email_time_str, subject, main_message)
        else :
            #change current sheet because email is of new month
            set_current_sheet(email_month_year, current_sheet)
            previous_sheet = current_sheet
            current_sheet = email_month_year
            #add new sheet for new month
            add_new_sheet(sheet, current_sheet, '1')
            #get sheet index of new sheet
            sheet_index = get_sheet_index(sheet, current_sheet)
            #add mail to new sheet
            add_mail_to_sheet(sheet, current_sheet, sheet_index, email_date_str, email_time_str, subject, main_message)
        #mark mail unread after putting mail in sheet using msg id 
        mark_unread(gmail_service, 'me', msg_Ids)
            
if __name__ == '__main__':   
    main()