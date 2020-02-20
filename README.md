Extracting data using SQL query into pandas dataframe. 
Dataframe contains user name and email. 
Sending email to each user with necessary data. 


Libraries used in this project:
import smtplib
import pyodbc
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import datetime as dt
