#!/usr/local/bin/python2.7

#####################################################################
# Digital Edge Informity weekly report script                       #
# httpd://www.digitaledge.net                                       #
#                                                                   #
# This script may not be published or re-used without permission.   #
#                                                                   #
# Email support@digitaledge.net for immediate assistance            #
#####################################################################

import ldap
import xlrd
import array
import datetime
import re
import subprocess
import xlutils
import paramiko
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

DBSTSRV = '10.227.9.205'
DBDBSRV = '10.227.9.219'
DBDB9ASRV = '10.227.9.223'
DBDBSUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb.sh'
DBDB9ASUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb9a.sh'
EXCELFILE = 'Z:\\04-03_Replication_failure_control_doc.xlsx'

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


def try_ldap():
    LDAP_SERVER_EMG = "ldap://10.227.9.233"
    BIND_DN = "cn=Manager,dc=infomity,dc=net"
    BIND_PASS = "secret"
    USER_BASE = "nettsInstitutionCode=2012595,o=institutions,dc=infomity,dc=net"
    searchFilter = '(objectclass=*)'
    attrs = ['sn']
    try:
        ldap.set_option(ldap.OPT_X_TLS_REQUIRE_CERT, 0)
        lcon_emg = ldap.initialize(LDAP_SERVER_EMG)
        lcon_emg.simple_bind_s(BIND_DN, BIND_PASS)
        ldap_result_id = lcon_emg.search_s(USER_BASE, ldap.SCOPE_SUBTREE, searchFilter, attrs)
        for dn,entry in ldap_result_id:
            print 'Processing',repr(dn)

    except ldap.LDAPError, e:
        print e

def get_new_tickest_from_db(server):
    i = 0
    no_space_count = 0
    ticket_list = []
    today = datetime.datetime.today()
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    cmd = 'cd /usr/local/pgsql/data/infomity-mainte/log/; cat `ls -tr infomity-userDBRepliDelayPingCSV-ForStatistics-' + str(today.year) + '* | tail -n 1` | egrep -iv "@@@@@|2000010|2008977|2004785|2008195|2009247|2009275|2009612" | grep ",1"'
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    for x in range(len(teststr)):
        str1 = re.split("\|", teststr[x])
        if len(str1) == 3:
            str2 = re.split(":", str1[2])
            if int(str2[0]) > 100:
                ticket = Ticket()
                ticket.customer_dbid = str1[1]
                ticket_list.append(ticket)
                i = i + 1
    return ticket_list

def remove_dublicates(list1, list2):
    i = 0
    listidx = []
    list_tmp = []
    for x in range(len(list1)):
        for y in range(len(list2)):
            if list1[x].customer_dbid == list2[y].customer_dbid:
                listidx.append(x)
                i += 1

    for x in range(len(list1)):
        flag = True
        for y in range(len(listidx)):
            if x == listidx[y]:
                flag = False
        if flag:
            list_tmp.append(list1[x])
	
    return list_tmp	

xl_workbook = xlrd.open_workbook(EXCELFILE)	
closed_dbdblist = de_read_sheet('Closed-dbdb', 3)
closed_dbdb9alist = de_read_sheet('Closed-dbdb9a', 3)
dbdblist = de_read_sheet('dbdb', 11)
dbdb9alist = de_read_sheet('dbdb9a', 11)

#try_ldap()
dbdb_new = get_new_tickest_from_db(DBDBSRV)
dbdb9a_new = get_new_tickest_from_db(DBDB9ASRV)
dbdb_new = remove_dublicates(dbdb_new, dbdblist)
dbdb9a_new = remove_dublicates(dbdb9a_new, dbdb9alist)
