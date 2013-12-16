#import sys
#import math
#import glob
import pandas
import matplotlib.pyplot as plt
import os
import csv
import zipfile
import datetime
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.Utils import COMMASPACE, formatdate
from email import Encoders

#Global Variables
mailServer = 'email-01'
sentFrom = 'jivory@bohunt.hants.sch.uk'
summaryReportRecipient = 'nleete@bohunt.hants.sch.uk'
allReportsRecipient = 'jivory@bohunt.hants.sch.uk'


def send_mail(send_from, send_to, subject, text, filename, server):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(filename, 'rb').read())
    Encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(filename))
    msg.attach(part)

    smtp = smtplib.SMTP(server)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()


def send_to_hod(billCode, filename):
    address_book = 'addresses.csv'
    addresses = csv.DictReader(open(address_book, 'rb'), delimiter=',')

    for line in addresses:
        if (line['billingCode'] == billCode) and (line['userName'] != ''):
            send_mail(sentFrom, line['userName'] + '@bohunt.hants.sch.uk', 'Printing Report',
                      'Attached is the monthly printing report for ' + billCode + line['userName'], filename,
                      "email-01")

# Create Reports directory
if not os.path.exists('reports'):
    os.mkdir('reports')

# Get last months date
today = datetime.date.today()
first = datetime.date(day=1, month=today.month, year=today.year)
lastMonth = first - datetime.timedelta(days=1)
datestamp = lastMonth.strftime("%Y-%m")

# Prepare Summary csv file for output
csvSummaryFile = open('reports/print_summary_' + str(datestamp) + '.csv', 'wb')
csvsummary = csv.writer(csvSummaryFile)

# Read in the source csv file
df = pandas.read_csv('allbilling.csv')

# Group the data by the column Group Name
grouped = df.groupby('cGroupName')

# Iterate through data by Group Name
for cGroupName, group in grouped:
    folder_name = cGroupName + '_' + str(datestamp)
    if not os.path.exists('reports/%s' % folder_name):
        os.mkdir('reports/%s' % folder_name)
        os.mkdir('reports/%s/source' % folder_name)
    csvDetailWrite = open('reports/%s/source/detail.csv' % folder_name, 'wb')
    group.to_csv(csvDetailWrite, float_format='%.2f')
    csvDetailWrite.close()
    summary = pandas.pivot_table(group, values='cAmount', rows='cUserWhoPrinted', aggfunc=sum)
    cvsBillSummaryWrite = open('reports/%s/source/summary.csv' % folder_name, 'wb')
    summary.to_csv(cvsBillSummaryWrite, float_format='%.2f')
    cvsBillSummaryWrite.close()
    summary.plot(kind='bar')
    plt.title(cGroupName + ' ' + str(datestamp))
    plt.ylabel('Money Spent')
    plt.xlabel('User Account')
    plt.subplots_adjust(bottom=0.3)
    plt.savefig('reports/%s/source/summary.png' % folder_name)
    plt.clf()

    csvBillSummaryRead = open('reports/%s/source/summary.csv' % folder_name, 'rb')
    totalReader = csv.reader(csvBillSummaryRead)
    totalSpent = 0
    for row in totalReader:
        totalSpent += float(row[1])
    csvBillSummaryRead.close()

    csvBillSummaryRead = open('reports/%s/source/summary.csv' % folder_name, 'rb')
    reader = csv.reader(csvBillSummaryRead)
    csvDetailRead = open('reports/%s/source/detail.csv' % folder_name, 'rb')
    reader2 = csv.reader(csvDetailRead)
    htmlfile = open('reports/%s/report.html' % folder_name, 'w')
    htmlfile.write('<center>')
    htmlfile.write('<h1>Printing Report for %s</h1>' % cGroupName)
    htmlfile.write('<hr>')
    htmlfile.write('<h2>Total Spend for ' + datestamp + '</h2>')
    htmlfile.write('<h3>' + chr(163) + str(totalSpent) + '</h3>')
    htmlfile.write('<hr>')
    htmlfile.write('<h2>Summary Chart</h2>')
    htmlfile.write('<img src="source/summary.png">')
    htmlfile.write('<h2>Summary Table</h2>')
    htmlfile.write('<table border="1">')
    htmlfile.write('<tr>')
    htmlfile.write('<th>User</th>')
    htmlfile.write('<th>Money Spent</th>')
    htmlfile.write('</tr>')
    for row in reader:
        colnum = 1
        htmlfile.write('<tr>')
        for column in row:
            htmlfile.write('<td>' + column + '</td>')
        htmlfile.write('</tr>')
    htmlfile.write('</table>')
    htmlfile.write('<hr>')
    htmlfile.write('<h2>Detail Table</h2>')
    htmlfile.write('<table border="1">')
    rownum = 0
    for row in reader2:
        if rownum == 0:
            htmlfile.write('<tr>')
            for column in row:
                htmlfile.write('<th>' + column + '</th>')
            htmlfile.write('</tr>')
        else:
            htmlfile.write('<tr>')
            for column in row:
                htmlfile.write('<td>' + column + '</td>')
            htmlfile.write('</tr>')
    htmlfile.write('</table>')
    htmlfile.write('</center>')
    htmlfile.close()
    csvsummary.writerow([cGroupName, totalSpent])
    csvDetailRead.close()
    csvBillSummaryRead.close()

    target_dir = 'reports/%s' % folder_name
    zip = zipfile.ZipFile('reports/' + cGroupName + '_PrintReport_' + str(datestamp) + '.zip', 'w',
                          zipfile.ZIP_DEFLATED)
    rootlen = len(target_dir) + 1
    for base, dirs, files in os.walk(target_dir):
        for file in files:
            fn = os.path.join(base, file)
            zip.write(fn, fn[rootlen:])
    zip.close()

    for root, dirs, files in os.walk('reports/' + folder_name, topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))
    os.rmdir('reports/' + folder_name)

    send_to_hod(cGroupName, 'reports/' + cGroupName + '_PrintReport_' + str(datestamp) + '.zip')

csvSummaryFile.close()

target_dir = 'reports'
zip = zipfile.ZipFile('All_PrintReports_' + str(datestamp) + '.zip', 'w', zipfile.ZIP_DEFLATED)
rootlen = len(target_dir) + 1
for base, dirs, files in os.walk(target_dir):
    for file in files:
        fn = os.path.join(base, file)
        zip.write(fn, fn[rootlen:])
zip.close()

send_mail(sentFrom, summaryReportRecipient, 'Printing Report Summary for ' + str(datestamp),
          'Here is the monthly Printing report summary for ' + str(datestamp),
          'reports/print_summary_' + str(datestamp) + '.csv', "email-01")
send_mail(sentFrom, allReportsRecipient, 'Printing Reports for ' + str(datestamp),
          'Here are the monthly Printing reports.  Enjoy!', 'All_PrintReports_' + str(datestamp) + '.zip', "email-01")
send_mail(sentFrom, 'jfroelich@bohunt.hants.sch.uk', 'Printing Reports for ' + str(datestamp),
          'Here are the monthly Printing reports.  Enjoy!', 'All_PrintReports_' + str(datestamp) + '.zip', "email-01")

for root, dirs, files in os.walk('reports', topdown=False):
    for name in files:
        os.remove(os.path.join(root, name))
    for name in dirs:
        os.rmdir(os.path.join(root, name))

os.rmdir('reports')
os.remove('allbilling.csv')
os.remove('All_PrintReports_' + str(datestamp) + '.zip')
