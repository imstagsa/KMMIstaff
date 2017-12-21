#####################################################################
# Digital Edge Informity weekly report script                       #
# httpd://www.digitaledge.net                                       #
#                                                                   #
# This script may not be published or re-used without permission.   #
# Last updated 12/18/2015                                           #
# Email support@digitaledge.net for immediate assistance            #
#####################################################################

import os
import re
import xlrd
import xlwt
import array
import datetime
import xlutils
import paramiko
import string
import smtplib
import requests
import de_classes
import email.message
from xlrd import open_workbook
from xlutils.copy import copy
from operator import attrgetter
from requests_ntlm import HttpNtlmAuth
from smtplib import SMTP_SSL as SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def read_filter_from_file(filename):
    f = open(filename, 'r')
    line = f.readline()
    f.close()
    return line

def write_log_file(filename):
    f = open(filename, 'r')
    line = f.readline()
    f.close()
    return line	
	
def exec_command_via_ssh(server, cmd):
    tmpstr = ""
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    stdin, stdout, stderr = ssh.exec_command(cmd)
    tmpstr = stdout.readlines()
    tmperr = stderr.readlines()
    ssh.close()
    return tmpstr

def mount_x():
    cmd = 'net use x: \"\\\\10.227.14.5\SharedDrive\ALL KMMI Personnel\Informity Collaboration\" /user:sevghenii kmmi@07122015'
    os.system(cmd)

def unmount_x():
    cmd = 'net use x: /delete /yes '
    os.system(cmd)

def str_to_date(workbook, value):
    try:
	    a1_tuple = xlrd.xldate_as_tuple(value, workbook.datemode)
	    d = datetime.datetime(*a1_tuple)
    except ValueError:
        d = str_to_date1(value)
    return d

#converts string 2015-12-31 to date
def str_to_date1(value):
    format = "%Y-%m-%d"
    try:
	    d = datetime.datetime.strptime(value, format)
    except ValueError,e:
        #print e
        d = str_to_date2(value)
    return d

#converts string 12/31/2015 to date
def str_to_date2(value):
    format = "%m/%d/%Y"
    try:
	    d = datetime.datetime.strptime(value, format)
    except ValueError,e:
        #print e
        d = str_to_date3(value)
    return d
	
#converts string 2015/12/31 to date
def str_to_date3(value):
    format = "%Y/%m/%d"
    try:
	    d = datetime.datetime.strptime(value, format)
    except ValueError,e:
        #print e
        d = str_to_date4(value)        
    return d

#converts string 20151231 to date
def str_to_date4(value):
    format = "%Y%m%d"
    try:
	    d = datetime.datetime.strptime(value, format)
    except ValueError,e:
        #print e
        d = datetime.datetime.today()
    return d

def de_read_sheet(workbook, sheet_name, indx):
    ticket_list = []
    xl_sheet = workbook.sheet_by_name(sheet_name)
    for x in xrange(indx, xl_sheet.nrows):    # Iterate through rows
        if len(str(xl_sheet.cell(x, 6).value)) > 0:
            ticket = de_classes.Ticket()
            ticket.ticket_number = xl_sheet.cell(x, 1).value
            if len(str(xl_sheet.cell(x, 2).value)) > 1:
               ticket.ticket_close_date = str_to_date(workbook, xl_sheet.cell_value(x, 2))
            ticket.ticket_open_date = str_to_date(workbook, xl_sheet.cell_value(x, 3))
            ticket.author = xl_sheet.cell(x, 4).value
            ticket.customer_name = xl_sheet.cell(x, 5).value
            ticket.customer_dbid = xl_sheet.cell(x, 6).value
            ticket.manual_synchronization = xl_sheet.cell(x, 7).value
            ticket.support_date_end = str_to_date(workbook, xl_sheet.cell_value(x, 8))
            ticket.contract_date_start = str_to_date(workbook, xl_sheet.cell_value(x, 9))
            ticket.another_db_exist = xl_sheet.cell(x, 10).value
            ticket.another_db_status = xl_sheet.cell(x, 11).value
            ticket.table_exist_check = xl_sheet.cell(x, 12).value
            ticket.last_db_access_date = str_to_date(workbook, xl_sheet.cell(x, 13).value)
            ticket.last_db_sync_date  = str_to_date(workbook, xl_sheet.cell(x, 14).value)
            ticket.log_check = xl_sheet.cell(x, 15).value
            ticket.st_server = xl_sheet.cell(x, 16).value
            ticket.kim_system_ini_timestamp = str_to_date(workbook, xl_sheet.cell_value(x, 17))
            ticket.image_pilot_version = str(xl_sheet.cell(x, 18).value)
            ticket.comment = xl_sheet.cell(x, 20).value
            ticket.action1_db = xl_sheet.cell(x, 21).value
            ticket.action1_request = xl_sheet.cell(x, 22).value
            ticket.comment_out_from_st = xl_sheet.cell(x, 23).value
            ticket.comment_out_from_pg_hba = str_to_date(workbook, xl_sheet.cell_value(x, 24))
            ticket.action3_request = xl_sheet.cell(x, 25).value
            ticket.call_center_comment = xl_sheet.cell(x, 26).value
            ticket.sheet_name = sheet_name
            ticket.databank_enabled = True
            ticket_list.append(ticket)
    return ticket_list
	
def de_read_sheet_comment_out(workbook, sheet_name, indx):
    ticket_list = []
    xl_sheet = workbook.sheet_by_name(sheet_name)
    for x in xrange(indx, xl_sheet.nrows):    # Iterate through rows
        if len(str(xl_sheet.cell(x, 6).value)) > 0:
            ticket = de_classes.Ticket()
            ticket.ticket_number = xl_sheet.cell(x, 1).value
            if len(str(xl_sheet.cell(x, 2).value)) > 1:
               ticket.ticket_close_date = str_to_date(workbook, xl_sheet.cell_value(x, 2))
            ticket.ticket_open_date = str_to_date(workbook, xl_sheet.cell_value(x, 3))
            ticket.author = xl_sheet.cell(x, 4).value
            ticket.customer_name = xl_sheet.cell(x, 5).value
            ticket.customer_dbid = xl_sheet.cell(x, 6).value
            ticket.image_pilot_version = str(xl_sheet.cell(x, 7).value)
            ticket.comment = xl_sheet.cell(x, 8).value
            ticket.sheet_name = sheet_name
            ticket_list.append(ticket)
    return ticket_list

def send_email(server, mailfrom, mailto, subject, body):
    msg = email.message.Message()
    msg['Subject'] = subject
    msg['From'] = mailfrom
    msg['To'] = ', '.join( mailto )
    msg.add_header('Content-Type','text/html')
    msg.set_payload(body)
    server = smtplib.SMTP(server)
    server.set_debuglevel(1)
    server.ehlo()
    server.starttls()
    server.ehlo()
    session = requests.Session()
    password = '8554476bdaf54455bff266b07107674b'
    #password = create_LM_hashed_password_v1(password)
    #print password.decode('hex')
    session.auth = HttpNtlmAuth('DIGITALEDGE\\esimacenco', password.decode('hex'), session)
    #server.login('esimacenco@digitaledge.net', 'password')
    server.sendmail(mailfrom, mailto, msg.as_string())
    server.quit()
    session.close()