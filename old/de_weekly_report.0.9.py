#!/usr/local/bin/python2.7

#####################################################################
# Digital Edge Informity weekly report script                       #
# httpd://www.digitaledge.net                                       #
#                                                                   #
# This script may not be published or re-used without permission.   #
#                                                                   #
# Email support@digitaledge.net for immediate assistance            #
#####################################################################

import xlrd
import array
import datetime
import re
import subprocess
import xlutils
import paramiko
import smtplib
import requests
from requests_ntlm import HttpNtlmAuth
from smtplib import SMTP_SSL as SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#SMTPSERVER='edge.digitaledge.net'
SMTPSERVER='208.74.200.21'
SENDFROM='esimacenco@digitaledge.net'
SENDTO=['esimacenco@digitaledge.net']
DBDBSRV = '10.227.9.219'
DBDB9ASRV = '10.227.9.223'
DBDBSUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb.sh'
DBDB9ASUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb9a.sh'

class Ticket:
    ticket_number = ''
    ticket_close_date = ''
    ticket_open_date = ''
    author = 'Evghenii' #4
    customer_name = ''
    customer_dbid = ''
    manual_synchronization = '' #May be OK or NG
    contract_date_end = ''
    contract_date_start = ''
    another_db_exist = ''
    another_db_status = ''
    table_exist_check = '' #OK, NULL or DOWN
    last_db_access_date = ''
    last_db_sync_date = ''
    log_check = ''
    st_server = 'ST1' #always ST1
    kim_system_ini_timestamp = ''
    image_pilot_version = ''
    comment = ''


filename = 'Z:\\04-03_Replication_failure_control_doc.xlsx'
xl_workbook = xlrd.open_workbook(filename)

def str_to_date(value):
    try:
	a1_tuple = xlrd.xldate_as_tuple(value, xl_workbook.datemode)
	d = datetime.datetime(*a1_tuple)
    except ValueError:
	d = str_to_date2(value)
    return d

#converts string 2015-10-08 to date
def str_to_date1(value):
    format = "%Y-%M-%d"
    try:
	    d = datetime.datetime.strptime(value, format)
    except ValueError:
	    d = datetime.datetime.today()
    return d

#converts string 12/31/2015 to date
def str_to_date2(value):
    format = "%M/%d/%Y"
    try:
	    d = datetime.datetime.strptime(value, format)
    except ValueError:
	#print value
	    d = datetime.datetime.today()
    return d	
	
def de_read_sheet(sheet_name, indx):
    i = 0
    ticket_list = []
    xl_sheet = xl_workbook.sheet_by_name(sheet_name)
    for x in xrange(indx, xl_sheet.nrows):    # Iterate through rows
	    ticket = Ticket()
	    ticket.ticket_number = xl_sheet.cell(x, 1)
	    ticket.ticket_close_date = str_to_date(xl_sheet.cell_value(x, 2))
	    ticket.ticket_open_date = str_to_date(xl_sheet.cell_value(x, 3))
	    ticket.author = xl_sheet.cell(x, 4)
	    ticket.customer_name = xl_sheet.cell(x, 5)
	    ticket.customer_dbid = xl_sheet.cell(x, 6).value
	    ticket.manual_synchronization = xl_sheet.cell(x, 7)
	    ticket.contract_date_end = str_to_date(xl_sheet.cell_value(x, 8))
	    ticket.contract_date_start = str_to_date(xl_sheet.cell_value(x, 9))
	    ticket.another_db_exist = xl_sheet.cell(x, 10)
	    ticket.another_db_status = xl_sheet.cell(x, 11)
	    ticket.table_exist_check = xl_sheet.cell(x, 12)
	#print i
	    ticket.last_db_access_date = str_to_date(xl_sheet.cell(x, 13).value)
	    ticket.last_db_sync_date  = str_to_date(xl_sheet.cell(x, 14).value)
	    ticket.log_check = xl_sheet.cell(x, 15)
	    ticket.st_server = xl_sheet.cell(x, 16)
	    ticket.kim_system_ini_timestamp = str_to_date(xl_sheet.cell_value(x, 17))
	    ticket.image_pilot_version = xl_sheet.cell(x, 18)
	    ticket.comment = xl_sheet.cell(x, 19)
	    ticket_list.append(ticket)
	    i = i + 1
    return ticket_list

def removed_from_list(listt):
    rem_count = 0
    today = datetime.datetime.today()
    dates = [today + datetime.timedelta(days=i) for i in range(0 - today.weekday(), 7 - today.weekday())]
    d1 = today + datetime.timedelta(0 - today.weekday())
    d2 = today + datetime.timedelta(7 - today.weekday())
    for x in xrange(0, listt.__len__()):
	    if d1 <  listt[x].ticket_close_date < d2:
	        rem_count += 1
    return rem_count

def added_to_list(listt1, listt2):
    rem_count = 0
    today = datetime.datetime.today()
    dates = [today + datetime.timedelta(days=i) for i in range(0 - today.weekday(), 7 - today.weekday())]
    d1 = today + datetime.timedelta(0 - today.weekday())
    d2 = today + datetime.timedelta(7 - today.weekday())
    for x in xrange(0, listt1.__len__()):
	    if d1 <  listt1[x].ticket_open_date < d2:
	        rem_count += 1
    for x in xrange(0, listt2.__len__()):
	    if d1 <  listt2[x].ticket_open_date < d2:
	        rem_count += 1
    return rem_count

def get_summary_script(server, scriptname):
    tmpstr = ""
    cmd="sh " + scriptname
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    stdin, stdout, stderr = ssh.exec_command(cmd)
    tmpstr = stdout.readlines()
    ssh.close()
    #print tmpstr
    return tmpstr

def check_no_free_space(server, listt):
    print server
    no_space_count = 0
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    for x in xrange(0, listt.__len__()):
	    sitestr = listt[x].customer_dbid[:7]
	    cmd="sh /infomity-mainte/bin/InstitutionInfoCheck/InstitutionInfoGet_US.sh "+ sitestr +" | grep \"EmptySize:0\[KB\]\""
	    stdin, stdout, stderr = ssh.exec_command(cmd)
	    teststr = stdout.readlines()
	    if len(teststr) > 0:
			no_space_count = no_space_count + 1
		
    ssh.close()
    return no_space_count

def check_contract_issue(server, listt):
    print server
    no_space_count = 0
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    for x in xrange(0, listt.__len__()):
	    sitestr = listt[x].customer_dbid[:7]
	    cmd="sh /infomity-mainte/bin/InstitutionInfoCheck/InstitutionInfoGet_US.sh "+ sitestr +" | grep \"Data Bank Service\" |grep 2020/01/01"
	    stdin, stdout, stderr = ssh.exec_command(cmd)
	    teststr = stdout.readlines()
	    if len(teststr) == 0:
			no_space_count = no_space_count + 1
			
    ssh.close()
    return no_space_count

def build_report_message():

    msg = MIMEMultipart('alternative')
    msg['Subject'] = "Weekly Report of Informity DB Connection Status"
    msg['From'] = SENDFROM
    msg['To'] = SENDTO
	
    tmpdb = get_summary_script(DBDBSRV, DBDBSUMMARY)
    tmpdb9a = get_summary_script(DBDB9ASRV, DBDB9ASUMMARY)
    
    #Create the body of the message (a plain-text and an HTML version).
    html = """\
    <html>
	<head>
	    <style><!--
		@font-face
		    {font-family:Calibri;
		    panose-1:2 15 5 2 2 2 4 3 2 4;}
		/* Style Definitions */
		p.MsoNormal, li.MsoNormal, div.MsoNormal
		    {margin:0in;
		    margin-bottom:.0001pt;
		    font-size:11.0pt;
		    font-family:"Calibri","sans-serif";}
		a:link, span.MsoHyperlink
		    {mso-style-priority:99;
		    color:blue;
		a:visited, span.MsoHyperlinkFollowed
		    {mso-style-priority:99;
		    color:purple;
		    text-decoration:underline;}
		span.EmailStyle17
		    {mso-style-type:personal-compose;
		    font-family:"Calibri","sans-serif";
		    color:windowtext;}
		.MsoChpDefault
		    {mso-style-type:export-only;
		    font-family:"Calibri","sans-serif";}
		@page WordSection1
		    {size:8.5in 11.0in;
		    margin:1.0in 1.0in 1.0in 1.0in;}
		div.WordSection1
		    {page:WordSection1;}--->
	    </style>
	</head>
	<body>
	<div class=WordSection1>
    	<p>ALL
    	    <br> Please see below for this week&#8217;s report on DB connection status:<br>
    	<p class=MsoNormal><span style='color:#1F497D'>*********** </span>
    	    <span style='color:#C55A11'>Hostname: DBDB - Today's Report Count Summary </span>
    	    <span style='color:#1F497D'>*****************<o:p></o:p></span>
    	</p>
    	"""
    html = html + "<p class=MsoNormal>" + tmpdb[2] + "</p>"
    html = html + "<p class=MsoNormal>" + tmpdb[3] + "</p>"
    html = html + "<p class=MsoNormal>" + tmpdb[4] + "</p>"
    html = html + "<p class=MsoNormal>" + tmpdb[5] + "</p>"
    html = html + "<p class=MsoNormal><span style='color:#C00000'>" +tmpdb[6] + "</span></p>"
    html = html + "<p class=MsoNormal>Sync NG No space = " + str(check_no_free_space(DBDBSRV, dbdblist)) + "</p>"
    html = html + "<p class=MsoNormal>Sync NG Over 100h Contract issue = " + str(check_contract_issue(DBDBSRV, dbdblist)) + "</p>"
    html = html + "<p class=MsoNormal>Removed from list = " + str(removed_from_list(closed_dbdblist)) + "</p>"
    html = html + "<p class=MsoNormal>Add-on to list = " + str(added_to_list(dbdblist,closed_dbdblist)) + "</p>"
    html = html + "</p></body></html>"
	
    html = html + """\
    <p><span style='color:#1F497D'>*********** </span>
    	    <span style='color:#C55A11'>Hostname: DBDB9a - Today's Report Count Summary  </span>
    	    <span style='color:#1F497D'>*****************<o:p></o:p></span>
    	</p>
    """
    html = html + "<p class=MsoNormal>" + tmpdb9a[2] + "</p>"
    html = html + "<p class=MsoNormal>" +tmpdb9a[3] + "</p>"
    html = html + "<p class=MsoNormal>" +tmpdb9a[4] + "</p>"
    html = html + "<p class=MsoNormal>" +tmpdb9a[5] + "</p>"
    html = html + "<p class=MsoNormal><span style='color:#C00000'>" +tmpdb9a[6] + "</span></p>"
    html = html + "<p class=MsoNormal>Sync NG No space = " + str(check_no_free_space(DBDB9ASRV, dbdb9alist)) + "</p>"
    html = html + "<p class=MsoNormal>Sync NG Over 100h Contract issue = " + str(check_contract_issue(DBDB9ASRV, dbdb9alist)) + "</p>"
    html = html + "<p class=MsoNormal>Removed from list = " + str(removed_from_list(closed_dbdb9alist)) + "</p>"
    html = html + "<p class=MsoNormal>Add-on to list = " + str(added_to_list(dbdb9alist,closed_dbdb9alist)) + "</p>"
    html = html + """\
    <p></p>
    <p class=MsoNormal>
	<b><span style='color:#002776'>Evghenii Simacenco</span></b>
	<span style='color:#1F497D'><o:p></o:p></span>
    </p>
    <p class=MsoNormal>
	<b><span style='font-size:9.0pt;color:#002776'>DIGITAL EDGE | System Administrator</span></b>
	<span style='color:#1F497D'><o:p></o:p></span>
    </p>
    <p class=MsoNormal>
	<span style='font-size:9.0pt;color:#000066'><a href=\"mailto:esimacenco@digitaledge.net\"><span style='color:#0563C1'>esimacenco@digitaledge.net</span></a></span><span style='color:#000066'><o:p></o:p></span></p><p class=MsoNormal><span style='font-size:9.0pt;color:#0070C0'>(718) 370-3353 x142<o:p></o:p></span></p><p class=MsoNormal><o:p>&nbsp;</o:p>
    </p>
    </div>
    </body>
    </html>"""

    part = MIMEText(html, 'html')
    msg.attach(part)
	
    #print msg
    return msg

def send_email(msg):
    server.connect(SMTPSERVER)
    server.set_debuglevel(1)
    server.ehlo()
    #server.starttls()
    server.ehlo()
    session = requests.Session()
    session.auth = HttpNtlmAuth('DIGITALEDGE\\esimacenco','logSBrawn45', session)
    server.sendmail(SENDFROM, SENDTO, msg.as_string())
    server.quit()
    session.close()

closed_dbdblist = de_read_sheet('Closed-dbdb', 3)
closed_dbdb9alist = de_read_sheet('Closed-dbdb9a', 3)
dbdblist = de_read_sheet('dbdb', 11)
dbdb9alist = de_read_sheet('dbdb9a', 11)
send_email(build_report_message())
