import pandas as pd
import numpy as np

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
        for staff in object_list:
            if staff.last_name == lname:
                return staff
        
    
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

def load_and_pickle(file):
	df = pd.read_excel(file)
	for row in df.iterrows():
		full_name = row[1][0].split(', ')
		fname = full_name[1].strip()
		lname = full_name[0].strip()
		email = row[1][1]

		


load_and_pickle('S:\\resident\\resident_emails.xlsx')
