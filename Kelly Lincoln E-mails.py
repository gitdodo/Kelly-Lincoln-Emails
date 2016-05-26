import pyodbc
from openpyxl import Workbook
from datetime import date, timedelta
import os
import win32com.client as win32

cnxn = pyodbc.connect('Trusted_Connection=yes', driver = '{SQL Server}',server = 'SQLSYTLSP.SJM.COM,1433', database = 'TechnicalServices')
cursor = cnxn.cursor()

cursor.execute("""
SELECT CallActivity.ID AS TSRefNum, CAST(CONVERT(VARCHAR, PRHistory.DateSubmitted, 120) AS DATETIME) as DateSubmitted, PRHistory.SubmissionType, PRHistory.PRSite, CallActivity.PrimaryDeviceType,
CallActivity.PrimaryModel, CallActivity.PrimarySerial, CallActivity.RegionID FROM PRHistory JOIN
ComplaintRecord on PRHistory.ComplaintRecordID = ComplaintRecord.ID JOIN
CallActivity on ComplaintRecord.ActivityID = CallActivity.ID
WHERE DateSubmitted > CONVERT(date, getdate()-1) and DateSubmitted <= CONVERT(date, getdate()+1) and EPIQEvents is null
Order by DateSubmitted 
"""
)

#get colnames from openpyxl
columns = [column[0] for column in cursor.description]    

#open workbook
wb = Workbook()
ws = wb.active

#append column names to header
ws.append(columns)

#append rows to 
for row in cursor:
    l = list(row)
    ws.append(l)

today = date.today()
today = today.strftime("%m.%d")

yesterday = date.today() - timedelta(days=1)
yesterday = yesterday.strftime("%m.%d")

filename = 'C:/Users/pengk02/Desktop/Reports Sent Kelly Lincoln/Reports Sent %s-%s.xlsx' %(yesterday,today)
wb.save(filename)
os.startfile(filename, 'open')
print (filename)
cnxn.close()

with open('C:/Users/pengk02/AppData/Roaming/Microsoft/Signatures/New SJM Standard 1.6.htm') as myfile:
    data=myfile.read()

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'klincoln@sjm.com; sy_complaints@sjm.com'
mail.CC = 'arothrock@sjm.com'
mail.Subject = 'Reports Sent %s-%s' %(yesterday,today)
mail.GetInspector 
mail.HtmlBody = '<body>Hi Kelly, <br> Please find included the reports sent from %s-%s.<br><br>Regards,</body>' %(yesterday,today) + data
mail.Attachments.Add(filename)

mail.Display(True)
