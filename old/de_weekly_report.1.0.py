#!/usr/local/bin/python2.7

#####################################################################
# Digital Edge Informity weekly report script                       #
# httpd://www.digitaledge.net                                       #
#                                                                   #
# This script may not be published or re-used without permission.   #
# Last updated 10/28/2015                                           #
# Email support@digitaledge.net for immediate assistance            #
#####################################################################

import re
import xlrd
import array
import datetime
import xlutils
import paramiko
import smtplib
import requests
import subprocess
import de_classes
from de_utils import de_read_sheet
import email.message
from requests_ntlm import HttpNtlmAuth
from smtplib import SMTP_SSL as SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

SMTPSERVER='edge.digitaledge.net'
SENDFROM='esimacenco@digitaledge.net'
SENDTO=['esimacenco@digitaledge.net']
DBDBSRV = '10.227.9.219'
DBDB9ASRV = '10.227.9.223'
DBDBSUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb.sh'
DBDB9ASUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb9a.sh'
#EXCELFILE = 'Z:\\04-03_Replication_failure_control_doc.xls'
EXCELFILE = 'app\\04-03_Replication_failure_control_doc.xls'

#reading one sheet from xls file and returns number of closed tickets during current week.
def removed_from_list(listt):
    rem_count = 0
    today = datetime.datetime.today()
    dates = [today + datetime.timedelta(days=i) for i in range(0 - today.weekday(), 7 - today.weekday())]
    d1 = today + datetime.timedelta(0 - today.weekday())
    d2 = today + datetime.timedelta(7 - today.weekday())
    for ticket in listt:
        if d1 <  ticket.ticket_close_date < d2:
            rem_count += 1
    return rem_count
	
#reading one sheet from xls file and returns number of added tickets during current week.
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
    return tmpstr

def check_no_free_space(server, tickets):
    print server
    no_space_count = 0
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    for ticket in tickets:
	    cmd="sh /infomity-mainte/bin/InstitutionInfoCheck/InstitutionInfoGet_US.sh "+ ticket.customer_dbid[:7] +" | grep \"EmptySize:0\[KB\]\""
	    stdin, stdout, stderr = ssh.exec_command(cmd)
	    teststr = stdout.readlines()
	    if len(teststr) > 0:
			no_space_count = no_space_count + 1
		
    ssh.close()
    return no_space_count

def check_contract_issue(server, tickets):
    print server
    contact_issues = 0
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    for ticket in tickets:
	    cmd="sh /infomity-mainte/bin/InstitutionInfoCheck/InstitutionInfoGet_US.sh "+ ticket.customer_dbid[:7] +" | grep \"Data Bank Service\" |grep \"2020/01/01\""
	    stdin, stdout, stderr = ssh.exec_command(cmd)
	    teststr = stdout.readlines()
	    if len(teststr) == 0:
			contact_issues = contact_issues + 1
			
    ssh.close()
    return contact_issues

def build_report_message():

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

    msg = email.message.Message()
    msg['Subject'] = "Weekly Report of Informity DB Connection Status"
    msg['From'] = SENDFROM
    msg['To'] = ', '.join( SENDTO )
    msg.add_header('Content-Type','text/html')
    msg.set_payload(html)
    return msg

def send_email(msg):
    server = smtplib.SMTP(SMTPSERVER)
    #server.set_debuglevel(1)
    server.ehlo()
    session = requests.Session()
    session.auth = HttpNtlmAuth('DIGITALEDGE\\esimacenco','logSBrawn45', session)
    server.sendmail(SENDFROM, SENDTO, msg.as_string())
    server.quit()
    session.close()

xl_workbook = xlrd.open_workbook(EXCELFILE)
closed_dbdblist = de_read_sheet(xl_workbook, 'Closed-dbdb', 3)
closed_dbdb9alist = de_read_sheet(xl_workbook, 'Closed-dbdb9a', 3)
dbdblist = de_read_sheet(xl_workbook, 'dbdb', 11)
dbdb9alist = de_read_sheet(xl_workbook, 'dbdb9a', 11)
send_email(build_report_message())
