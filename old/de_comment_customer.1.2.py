#!/usr/local/bin/python2.7

#####################################################################
# Digital Edge Informity weekly report script                       #
# httpd://www.digitaledge.net                                       #
#                                                                   #
# This script may not be published or re-used without permission.   #
# Last updated 11/04/2015                                           #
# Email support@digitaledge.net for immediate assistance            #
#####################################################################

import re
import os
import ldap
import xlrd
import xlwt
import array
import shutil
import datetime
import subprocess
import xlutils
import paramiko
import smtplib
import string
import helper_ldap
import de_classes
from de_utils import de_read_sheet
from de_utils import str_to_date1
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from xlrd import open_workbook
from xlutils.copy import copy
from xlutils.margins import number_of_good_rows
from operator import attrgetter
from operator import itemgetter

#dbdb:/usr/local/pgsql/data/pg_hba.conf
#dbdb9a:/usr/local/pgsql/data/pg_hba.conf
#dbst:/infomity-mainte/bin/stcheck/SiteFilter.txt
#dbst:/Admin/scripts/site_filter.txt
#KMMIQADMIN
#net use z: "\\kmmi2\SharedDrive\ALL KMMI Personnel\Informity Collaboration" /user:sevghenii kmmi@10052015
#net use /delete ""\\kmmi2\SharedDrive\ALL KMMI Personnel\Informity Collaboration"
#import os
#cmd = 'net use z: "\\kmmi2\SharedDrive\ALL KMMI Personnel\Informity Collaboration" /user:sevghenii kmmi@10052015'
#os.system(cmd)

DBSTSRV = '10.227.9.205'
DBDBSRV = '10.227.9.219'
DBDB9ASRV = '10.227.9.223'
LDAPSRV = "ldap://10.227.9.233"
DBDBSUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb.sh'
DBDB9ASUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb9a.sh'
DBST_FILTER='app\\dbdst_filter.txt'
DBDB_FILTER='app\\dbdb_filter.txt'
DBDB9A_FILTER='app\\dbdb9a_filter.txt'

def comment_in_pg_hba(server, ticket):
    file_name='/Admin/scripts/pg_hba.conf'
    today = datetime.datetime.today()
    file_name_bkp=file_name+"_"+today.strftime('%m%d%Y')
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    cmd = 'cp '+file_name+' '+file_name_bkp
    stdin, stdout, stderr = ssh.exec_command(cmd)
    cmd = 'sed -i  \'s/^hostssl\ '+ticket.customer_dbid+'/#hostssl\ '+ticket.customer_dbid+'/\' /Admin/scripts/pg_hba.conf'
    stdin, stdout, stderr = ssh.exec_command(cmd)
    ssh.close()

def add_to_filter_in_dbst(server, ticket):
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    #
    cmd = 'str=`cat /Admin/scripts/SiteFilter.txt`; echo $str"|'+ticket.customer_dbid[:7]+','+ticket.customer_dbid[7:]+'" > /Admin/scripts/SiteFilter.txt'
    stdin, stdout, stderr = ssh.exec_command(cmd)
    cmd = 'echo "'+ticket.customer_dbid[:7]+','+ticket.customer_dbid[7:]+'" >> /Admin/scripts/site_filter.txt'
    stdin, stdout, stderr = ssh.exec_command(cmd)
    ssh.close()

def add_to_local_filter_file(filter_file, ticket):
    f = open('app\\workfile.txt', 'r')
    line = f.readline()
    #print line
    f.close()
    f = open('app\\workfile2.txt', 'w')
    line = line + "|"+ticket.customer_dbid[:7]+","+ticket.customer_dbid[7:]
    f.write(line)
    f.close()
    os.remove('app\\workfile.txt')
    os.rename('app\\workfile2.txt', 'app\\workfile.txt')

ticket = de_classes.Ticket()
ticket.customer_dbid = "urrrraaasdjfjhks"
add_to_filter_in_dbst(DBSTSRV, ticket)		
comment_in_pg_hba(DBSTSRV, ticket)	
add_to_local_filter_file(DBST_FILTER, ticket)

