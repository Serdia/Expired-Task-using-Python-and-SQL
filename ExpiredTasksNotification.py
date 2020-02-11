
import smtplib
import pyodbc
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import datetime as dt

connection_string = pyodbc.connect('DRIVER={SQL Server};server=MyServer;DATABASE=MyDB;Trusted_Connection=yes;')

# sql part
sql_query = """
select
	tblUsers.Name_FirstLast as AssignedTo
	,tblQuotes.ControlNo
	,tblUsers.EmailAddress as Email
    ,lstNoteTypes.Description as NoteType
	,tblNoteStore.Subject
	,CAST(tblNoteStore.CreatedDate as DATE)  as CreatedDate
	,CAST(tblNoteDiaries.DueDate as DATE) as DueDate
	,CAST(tblNoteDiaries.FinalDeadline as DATE) as FinalDeadline
FROM tblUsers 
LEFT JOIN tblNoteRecipients ON tblNoteRecipients.UserGUID = tblUsers.Userguid 
LEFT JOIN tblNoteEntries    ON tblNoteEntries.ID = tblNoteRecipients.EntryGUID --   Body
LEFT JOIN tblNoteStore      ON tblNoteEntries.NoteGUID = tblNoteStore.ID -- Subject 
LEFT JOIN dbo.lstNoteTypes  ON dbo.tblNoteStore.Type = dbo.lstNoteTypes.NoteTypeID
LEFT JOIN tblNoteEntities   ON tblNoteStore.ID = tblNoteEntities.NoteGUID
LEFT JOIN tblQuotes         ON tblQuotes.controlguid = dbo.tblNoteEntities.ControlGuid
LEFT JOIN tblNoteDiaries    ON dbo.tblNoteEntries.ID = dbo.tblNoteDiaries.EntryGUID  -- DueDate, FinalDeadline
WHERE (
            tblQuotes.QuoteID =
               (
                  SELECT MAX(tq1.QuoteID)
                  FROM  tblQuotes tq1
                  WHERE tblQuotes.ControlGuid = tq1.ControlGuid
               ) OR tblQuotes.QuoteID IS NULL
         )
	AND CAST(tblNoteDiaries.DueDate as DATE) = CAST(GETDATE() -0 AS DATE)
	AND CAST(dbo.tblNoteRecipients.CompletedDate AS DATE) IS NULL
	AND tblUsers.EmailAddress IS NOT NULL
	AND tblQuotes.CostCenterID IN (112,97,99,127,134) 
ORDER BY ControlNo
    """
df = pd.read_sql_query(sql_query, connection_string)

yesterday = dt.datetime.now() - dt.timedelta(days=1)
date = "'" + yesterday.strftime('%m-%d-%Y') + "'"

for email, d in df.groupby('Email'):
    # getting a total task number to use it in email subject 
    task_count = len(df.index)
    # selecting all rows for each email and saving it to df1
    df1 = df.loc[df['Email'] == email] # select * from df where email = email
    # removing column 'Email' from df1 because user does not need to see its own email
    df1 = df1.loc[:, df.columns != 'Email']

    # using css to create pretty table with data
    message_start = f"""
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <h2>Report for " + date + "</h2>
    <title>Report for {date}</title>"""
    message_style = """
    <style type="text/css" media="screen">
        #customers {
        font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
        font-size: 12px;
        border-collapse: collapse;
        width: 100%;
        }

        #customers td, #customers th {
        border: 1px solid #ddd;
        padding: 8px;
        }

        #customers tr:nth-child(even){background-color: #f2f2f2;}

        #customers tr:hover {background-color: #ddd;}

        #customers th {
        padding-top: 12px;
        padding-bottom: 12px;
        text-align: left;
        background-color: #4CAF50;
        color: white;
        }
    </style>
    </head>
    <body>
    """
    email = 'oserdyuk@aligngeneral.com'
    user = 'oserdyuk@aligngeneral.com'
    subject = 'Test Subject'

    message_body = df1.to_html(index=False, table_id="customers") #set table_id to your css style name
    message_end = """</body>"""
    messages = (message_start + message_style + message_body + message_end)
    part = MIMEText(messages, 'html')  # create MIMEText
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = user
    msg['To'] = email
    msg.attach(part)  
    server = smtplib.SMTP('1.1.1.1: 25')
    server.starttls()
    server.sendmail(user, msg['To'], msg.as_string())
    server.close()
    print("Mail sent succesfully!")