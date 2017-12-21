#####################################################################
# Digital Edge Informity find and resolve customers who didn't      #
# synchronized  last 4 days                                         #
# httpd://www.digitaledge.net                                       #
#                                                                   #
# This script may not be published or re-used without permission.   #
# Last updated 12/04/2015                                           #
# Email support@digitaledge.net for immediate assistance            #
#####################################################################

import re
import os
import ldap
import xlrd
import xlwt
import array
import datetime
import subprocess
import xlutils
import paramiko
import smtplib
import string
import helper_ldap
import de_classes
from de_utils import mount_x
from de_utils import unmount_x
from de_utils import send_email
from de_utils import str_to_date1
from de_utils import de_read_sheet
from de_utils import exec_command_via_ssh
from de_utils import read_filter_from_file
from de_utils import de_read_sheet_comment_out
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from xlrd import open_workbook
from xlutils.copy import copy
from xlutils.margins import number_of_good_rows
from operator import attrgetter
from operator import itemgetter

SMTPSERVER='mc1.digitaledge.net'
SENDFROM='esimacenco@digitaledge.net'
SENDTO=['esimacenco@digitaledge.net']
DBSTSRV = '10.227.9.205'
DBDBSRV = '10.227.9.219'
DBDB9ASRV = '10.227.9.223'
LDAPSRV = "ldap://10.227.9.233"
DBDBSUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb.sh'
DBDB9ASUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb9a.sh'
#EXCELFILE = 'app\\04-03_Replication_failure_control_doc_11.19.2015.xls'
#EXCELFILE2 = 'app\\04-03_Replication_failure_control_doc3.xls'
EXCELFILE = 'X:\\04-03_Replication_failure_control_doc.xls'
EXCELFILE2 = 'X:\\04-03_Replication_failure_control_doc2.xls'
d = datetime.datetime.today()
EXCEFILELBKP = 'X:\\Other_Docs\\failure_control_docs\\04-03_Replication_failure_control_doc_'+str(d.strftime('%m.%d.%Y'))+'.xls'
DBST_FILTER='app\\dbst_filter.txt'
DBDB_FILTER='app\\dbdb_filter.txt'
DBDB9A_FILTER='app\\dbdb9a_filter.txt'

def exec_ldap_query(server, basedn , searchFilter, base):
    values = []
    BIND_DN = "cn=Manager,dc=infomity,dc=net"
    BIND_PASS = "secret"
    attributeFilter = [base]
    try:
        ldap.set_option(ldap.OPT_X_TLS_REQUIRE_CERT, 0)
        lcon_emg = ldap.initialize(server)
        lcon_emg.simple_bind_s(BIND_DN, BIND_PASS)
        ldap_result_id = lcon_emg.search_s(basedn, ldap.SCOPE_SUBTREE, searchFilter, attributeFilter)
        res = helper_ldap.get_search_results(ldap_result_id)
        for i in res:
            #print i
            str = i.pretty_print2()
            if str is not None and len(str) > 0:
                values.append(str)
    except ldap.LDAPError, e:
        pass	
    return values

def fill_new_tickets(ticket):
    services_list = []

    SERVICE_CONTRACTS = "o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
    products = exec_ldap_query(LDAPSRV, SERVICE_CONTRACTS, "nettsInstitutionCode=*", "productNumber")
    lastProduct = ""
    summary_size = 0
    for product in products:
        SERVICE_CONTRACTS = "o=devices,productNumber=" + product + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        services = exec_ldap_query(LDAPSRV, SERVICE_CONTRACTS, "productNumber=*", "serialNumber")

        if str(product) != str(lastProduct):
            for serv in services:
                SERVICE_CONTRACTS = "cn=backupFullSize,serialNumber=" + serv + ",o=devices,productNumber=" + product + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
                full_size = exec_ldap_query(LDAPSRV, SERVICE_CONTRACTS, "cn=*", "numericValue")
                print product + " = " + serv + " = " + str(full_size)
                summary_size = summary_size + int(str(full_size[0]))
                print "Summary: " + str(summary_size)
        lastProduct = product
	
	
	
    return ticket

def copy_file(filesrc, filedst):
	cmd = "copy "+ filesrc +"   " +filedst
	os.system(cmd)
	
def print_date(function_name):
    d = datetime.datetime.today()
    print str(d) + " " + function_name

def print_list(tickets):
    print "List size"
    print len(tickets)
    for ticket in tickets:
        print ticket.customer_dbid

		
def read_from_filter_file():
    filemane='app\\site_filter_filled2.txt'
    f2 = open(filemane, 'r')
    ticket_list = []
    line = f2.readline()
    while line:
        ticket = de_classes.Ticket()
        ticket.customer_dbid = str(line)
        print ticket.customer_dbid
        ticket = fill_new_tickets(ticket)
        print ticket.customer_name
        ticket_list.append(ticket)
        for service in ticket.services:
            if service.service_name == 'DATABANK':
                f2.write(ticket.customer_name +"    "+ticket.customer_dbid + " " + service.service_name + " Size:" + service.full_size)
        line = f2.readline()
		
    f2.close()
    return ticket_list
	
print_date('begin')
try:
    read_from_filter_file()
    """ticket = de_classes.Ticket()
    ticket.customer_dbid = "20097905902330147823452A"
    ticket = fill_new_tickets(ticket)
    print ticket.customer_name
    for service in ticket.services:
        if service.service_name == 'DATABANK':
            print ticket.customer_name +"    "+ticket.customer_dbid + " " + service.service_name + " Size:" + service.full_size
	"""
except ValueError, e:
    print e

