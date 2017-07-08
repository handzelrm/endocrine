import pandas as pd
import numpy as np
import pickle
from datetime import datetime
from datetime import timedelta
import win32com.client as win32 #used to use local outlook program to send email
import smtplib #used to send emails from gmail

class Staff(object):
    """Creates an abstract base class
    
    Attributes:
        first_name: first name
        last_name: last name
        postition: position ie resident
        upmc_email: upmc email
    """
    gmail = None
    year = None    
    
    def __init__(self,first_name,last_name,upmc_email):
        self.first_name = first_name
        self.last_name = last_name
        self.upmc_email = upmc_email
    
    def __repr__(self):
        return "Staff('{}', '{}', '{}')".format(self.first_name,self.last_name,self.upmc_email)
    
    def __str__(self):
        return '{} {}'.format(self.first_name,self.last_name)
        
    @staticmethod
    def find_lastname(object_list,lname):
        if object_list is dict:
            for staff in object_list:
                if staff == lname:
                    return staff
        elif object_list is list:
            for staff in object_list:
                if staff.last_name == lname:
                    return staff
        else:
            print('error: needs to be list or dict')

    # @staticmethod
    # def find_lastdict(object_dict,lname):
    #     for staff in object_dict:
    #         if staff == lname:
    #             return staff
    # @staticmethod
    # def find_fullname(object_list,full_name):
    #     full_name = full_name.replace(' ','')
    #     full_name_list = full_name.split(',')
    #     fname = full_name_list[0]
    #     lname = full_name_list[1]
    #     for staff in object_list:
    #         print(staff)
    #         if staff.last_name == lname:
    #             if staff.first_name == fname:
    #                 return staff
        
    
class Resident(Staff):
    """Creates an abstract base class
    
    Attributes:
        first_name: first name
        last_name: last name
        postition: position ie resident
        upmc_email: upmc email
    """        
    def __init__(self,first_name,last_name,upmc_email,year=None,daysoff=None,number_of_calls=None): 
        super().__init__(first_name,last_name,upmc_email)
        self.position = 'Resident'
        self.year = year
        self.daysoff = []
        
    def __repr__(self):
        return "Staff('{}', '{}', '{}' {})".format(self.first_name,self.last_name,self.upmc_email,self.year)

    def add_dayoff(self,date):
        self.daysoff.append(date)
        
    def number_of_calls(self,number_of_calls):
        self.number_of_calls = number_of_calls

def load_and_pickle_res(path,file,sheet):
    res_dict = {}
    df = pd.read_excel(path+file,sheetname=sheet)
    df = df[['Resident Names (Last, First)','Resident Institutional Email']]
    df = df.dropna()
    
    for row in df.iterrows():
        full_name = row[1][0].split(', ')
        fname = full_name[1].strip()
        lname = full_name[0].strip()
        email = row[1][1]
        res_dict[lname] = Resident(fname,lname,email)

    with open(path+'res_dict.pickle','wb') as f:
        pickle.dump(res_dict,f)


def sendOutlookEmail(to):
    outlook = win32.Dispatch('Outlook.Application') #creates outlook applicatoin object
    mail = outlook.CreateItem(0) #creates outlook item mail
    mail.To = to #adds who to send it to
    mail.Subject = 'Endocrine Surgery Evaluation Reminder' #adds subject
    mail.Body = 'Hi,\n\nThis is a reminder to ask the attendings to fill out surgical evaluations after thyroidectomies and parathyroidectomies. Here are the links to the evals for your reference:\nThyroid: https://docs.google.com/forms/d/e/1FAIpQLScIJrSzFypVAivYRElGbVBssIpYwQAgzyCP9C4Bk5vOhgIBIw/viewform\nParathyroid: https://docs.google.com/forms/d/e/1FAIpQLSeZA7J6q5AneCX-8suob7omhPMRes82nIi3aA1IanKmKoYURg/viewform\n\nIf you need access to your evaluations please send me a gmail address that can be linked. Let me know if you have any questions or issues\n\nThanks,\nRob' #adds body
    mail.Send() #sends message

def createGmailServer():
    path = '/home/pi/Documents/python/'
    with open(path+'gm.pickle','rb') as f:
        gmun, gmpw = pickle.load(f)
    server = smtplib.SMTP_SSL('smtp.gmail.com',465)
    server.ehlo()
    server.login(gmun, gmpw)
    return server

def sendGmail(server,to):
    subject = 'Endocrine Surgery Evaluation Reminder' #adds subject
    body = 'Hi,\n\nThis is a reminder to ask the attendings to fill out surgical evaluations after thyroidectomies and parathyroidectomies. Here are the links to the evals for your reference:\nThyroid: https://docs.google.com/forms/d/e/1FAIpQLScIJrSzFypVAivYRElGbVBssIpYwQAgzyCP9C4Bk5vOhgIBIw/viewform\nParathyroid: https://docs.google.com/forms/d/e/1FAIpQLSeZA7J6q5AneCX-8suob7omhPMRes82nIi3aA1IanKmKoYURg/viewform\n\nIf you need access to your evaluations please send me a gmail address that can be linked. Let me know if you have any questions or issues\n\nThanks,\nRob' #adds body
    message = 'Subject: {}\n{}'.format(subject,body)
    server.sendmail('Robert.M.Handzel@gmail.com',to,message)


def check_if_notified(path,file):
    df = pd.read_excel(path+file)
    print(df.head())

    today = datetime.date.today()
    print(today)


    # df_pgy2 = pd.read_excel(path+file,sheetname='PGY 2')
    # # df_pgy2 = df_pgy2.transpose()
    # df_pgy5 = pd.read_excel(path+file,sheetname='PGY 4 & 5')



    # for col in df_pgy5.columns:
    #   df_row = df_pgy5.loc[df_pgy5[col]=='Endocrine']


def get_pgy1(path,file,sheet,send):
    df = pd.read_excel(path+file, sheetname=sheet,skiprows=3)
    df = df[df['PGY 1 Interns'].notnull()]
    df = df.drop(df.columns[[1,2]],axis=1)

    for cols in df.columns:
        if cols not in ['PGY 1 Interns']:
            date_span = cols
            date_list = date_span.split('-')
            date_list[0] = date_list[0].strip()
            date_list[1] = date_list[1].strip()
            start_date = datetime.strptime(date_list[0],'%m/%d/%y')
            end_date = datetime.strptime(date_list[1],'%m/%d/%y')

            # print('')
            # print('start date: {}'.format(start_date))
            # print('end date: {}'.format(end_date))

            if start_date-timedelta(days=1) <= datetime.now() <= end_date:
                current_col = cols
                # print(datetime.now()+timedelta(days=x_days))
                break

    df_current = df[['PGY 1 Interns',current_col]]
    df_endo = df_current[df_current[current_col]=='ENDO']

    print('pgy1:',end='\t')
    if df_endo.shape[0] == 0:
        print('No matches')
    elif df_endo.shape[0] == 1:
        resident = df_endo['PGY 1 Interns'].values[0]
        email = get_email(resident)
        print(resident)
        print('\t\tEmail:{}'.format(email))
        if send == 'windows':
            print('sending out outlook emails...')
            sendOutlookEmail(email)
        elif send == 'linux':
            print('sending out gmail emails...')
            server = createGmailServer()
            sendGmail(server=server,to='handzelrm@upmc.edu'):            
        else:
            print('no emails sent')
    else:
        print('More than one match')

def get_pgy2(path,file,sheet,send):
    df = pd.read_excel(path+file,sheetname=sheet,skiprows=3)
    # df = df.iloc[:21,:]
    # df = df[df.dropna()]
    df = df[df.NAME.notnull()]
    df = df.drop(df.columns[[1,2]],axis=1)
    # print(df)
    for cols in df.columns:
        if cols not in ['NAME','Unnamed: 14']:
            col_list = cols.split(' ')
            date_span = col_list[0]
            date_list = date_span.split('-')
            start_date = datetime.strptime(date_list[0],'%m/%d/%y')
            end_date = datetime.strptime(date_list[1],'%m/%d/%y')

            # print('')
            # print('start date: {}'.format(start_date))
            # print('end date: {}'.format(end_date))

            if start_date-timedelta(days=1) <= datetime.now() <= end_date:
                current_col = cols
                break

    df_current = df[['NAME',current_col]]
    df_endo = df_current[df_current[current_col]=='Endocrine']

    print('pgy2:',end='\t')
    if df_endo.shape[0] == 0:
        print('No matches')
    elif df_endo.shape[0] == 1:
        resident = df_endo.NAME.values[0]
        email = get_email(resident)
        print(resident)
        print('\t\tEmail:{}'.format(email))
        if send == 'windows':
            print('sending out outlook emails...')
            sendOutlookEmail(email)
        elif send == 'linux':
            print('sending out gmail emails...')
            server = createGmailServer()
            sendGmail(server=server,to='handzelrm@upmc.edu'):            
        else:
            print('no emails sent')
    else:
        print('More than one match')

def get_pgy5(path,file,sheet,send):
    df = pd.read_excel(path+file,sheetname=sheet,skiprows=19)
    df = df[df.RESIDENT.notnull()]

    for cols in df.columns:
        if cols not in ['RESIDENT','LR/SR Pager #s']:
            col_list = cols.split(' ')
            date_span = col_list[0]
            date_list = date_span.split('-')
            start_date = datetime.strptime(date_list[0],'%m/%d/%y')
            end_date = datetime.strptime(date_list[1],'%m/%d/%y')

            # print('')
            # print('start date: {}'.format(start_date))
            # print('end date: {}'.format(end_date))

            if start_date-timedelta(days=1) <= datetime.now() <= end_date:
                current_col = cols
                break

    df_current = df[['RESIDENT',current_col]]
    df_endo = df_current[df_current[current_col]=='Endocrine']

    print('pgy5:',end='\t')
    if df_endo.shape[0] == 0:
        print('No matches')
    elif df_endo.shape[0] == 1:
        resident = df_endo.RESIDENT.values[0]
        email = get_email(resident)
        print(resident)
        print('\t\tEmail:{}'.format(email))
        if send == 'windows':
            print('sending out outlook emails...')
            sendOutlookEmail(email)
        elif send == 'linux':
            print('sending out gmail emails...')
            server = createGmailServer()
            sendGmail(server=server,to='handzelrm@upmc.edu'):            
        else:
            print('no emails sent')
    else:
        print('More than one match')

def get_email(resident):
    path = 'S:/evals/'
    file = 'res_dict.pickle'
    lname = resident.split(',')[0]
    with open(path+file,'rb') as f:
        res_obj_dict = pickle.load(f)
    return res_obj_dict[lname].upmc_email


windows_path = 'S:/evals/'
linux_path = '/home/pi/Documents/python/'


send=False

if send == 'windows':
    path = windows_path
elif send == 'linux':
    path = 'linux_path'
else:
    path = windows_path #testing


load_and_pickle_res(path,'Master Spreadsheet.xlsx',sheet='Admin')
#note i deleted uncovered from column a for logic to prevent hardcoding

get_pgy1(path=path,file='gen_surg_schedule.xls',sheet='PGY 1',send=send)
get_pgy2(path=path,file='gen_surg_schedule.xls',sheet='PGY 2',send=send)
get_pgy5(path=path,file='gen_surg_schedule.xls',sheet='PGY 4 & 5',send=send)


# check_if_notified('S:\\resident\\','resident_endocrine_dates.xlsx')
