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

DBSTSRV = '10.227.9.205'
DBDBSRV = '10.227.9.219'
DBDB9ASRV = '10.227.9.223'
LDAPSRV = "ldap://10.227.9.233"
DBDBSUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb.sh'
DBDB9ASUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb9a.sh'
EXCELFILE = 'app\\04-03_Replication_failure_control_doc.xls'
EXCELFILE2 = 'app\\04-03_Replication_failure_control_doc2.xls'
#EXCELFILE = 'Z:\\04-03_Replication_failure_control_doc.xls'
#EXCELFILE2 = 'Z:\\04-03_Replication_failure_control_doc2.xls'
DBST_FILTER='app\\dbdst_filter.txt'
DBDB_FILTER='app\\dbdb_filter.txt'
DBDB9A_FILTER='app\\dbdb9a_filter.txt'

def read_filter_from_file(filename):
    f = open(filename, 'r')
    line = f.readline()
    #print line
    f.close()
    return line

def get_new_tickest_from_st(server, filter_file):
    ticket_list = []
    today = datetime.datetime.today()
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    exclude_list = read_filter_from_file(filter_file)
    cmd = 'sh /infomity-mainte/bin/stcheck/KimSystemini_chk.sh | egrep -iv \"'+exclude_list+'\" '
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    for x in range(len(teststr)):
        str1 = re.split(",", teststr[x])
        str2 = string.split(str1[2])
        kmidate = str_to_date1(str2[0])
        today = datetime.datetime.now()
        today = str_to_date1(str(today.year)+"/"+str(today.month)+"/"+str(today.day))
        diff_days = abs((today - kmidate).days)
        if diff_days > 4:
                ticket = de_classes.Ticket()
                ticket.customer_dbid = str1[0] + str1[1]
                str3 = str1[3]
                if str3[:3] == "1.7":
                    ticket.sheet_name = "dbdb9a"
                else:
                    ticket.sheet_name = "dbdb"
                ticket_list.append(ticket)
    ssh.close()
    return ticket_list
		
def get_new_tickest_from_db(server, filter_file):
    ticket_list = []
    today = datetime.datetime.today()
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    exclude_list = read_filter_from_file(filter_file)
    cmd = 'cd /usr/local/pgsql/data/infomity-mainte/log/; cat `ls -tr infomity-userDBRepliDelayPingCSV-ForStatistics-' + str(today.year) + '* | tail -n 1` | egrep -iv \"'+exclude_list+'\" | grep ",[0-1]"'
    print cmd
    stdin, stdout, stderr = ssh.exec_command(cmd)
    testserr = stderr.readlines()
    print testserr
    teststr = stdout.readlines()
    for x in range(len(teststr)):
        str1 = re.split("\|", teststr[x])
        if len(str1) == 3:
            str2 = re.split(":", str1[2])
            if int(str2[0]) > 100:
                ticket = de_classes.Ticket()
                ticket.customer_dbid = str1[1]
                if server == "10.227.9.223":
                    ticket.sheet_name = "dbdb9a"
                else:
                    ticket.sheet_name = "dbdb"
                ticket_list.append(ticket)
    ssh.close()
    return ticket_list

#Compair two arrays. Return new array with elements from first array which are not in the second array.
def remove_dublicates(list1, list2):
    listidx = []
    list_tmp = []
    for x in range(len(list1)):
        for y in range(len(list2)):
            if list1[x].customer_dbid == list2[y].customer_dbid:
                listidx.append(x)

    for x in range(len(list1)):
        flag = True
        for y in range(len(listidx)):
            if x == listidx[y]:
                flag = False
        if flag:
            list_tmp.append(list1[x])

    return list_tmp	
	
#return one ticket list with elements from firsts and second lists
def assembly_two_list(list1, list2):
    list_tmp = []
    for x in range(len(list1)):
        list_tmp.append(list1[x])
    for x in range(len(list2)):
        list_tmp.append(list2[x])
    return list_tmp

def fill_services_list(ticket):
    services_list = []
    SERVICE_CONTRACTS = "o=contracts,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
    services = exec_ldap_query(LDAPSRV, SERVICE_CONTRACTS, "(contractFlag=TRUE)", "infomityServiceCode")

    for serv in services:
        service = de_classes.Service()
        str1 = "infomityServiceCode=" + serv + ",o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "infomityServiceName")
        if len(str2) > 0:
            service.service_name = str2[0]
        
        str1 = "infomityServiceCode=" + serv + ",o=contracts,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "contractConclusionDay")
        if len(str2) > 0:
            service.contract_start = str2[0]

        str1 = "infomityServiceCode=" + serv + ",o=contracts,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "contractFinishDay")
        if len(str2) > 0:
            service.contract_end = str_to_date1(str2[0][:8])

        str1 = "infomityServiceCode=" + serv + ",o=contracts,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "infomityServiceStart")
        if len(str2) > 0:
            service.service_start = str2[0]

        str1 = "infomityServiceCode=" + serv + ",o=contracts,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "infomityServiceEnd")
        if len(str2) > 0:
            service.service_end = str2[0]

        str1 = "infomityServiceCode=" + serv + ",o=contracts,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "contractFlag")
        if len(str2) > 0:
            service.contract_flag = str2[0]
		
        str1 = "cn=InstCapacity,o=settings,infomityServiceCode=MGBOX,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "maximum")
        if len(str2) > 0:
            service.full_size = str2[0]

        str1 = "cn=mboxEmptySize,o=settings,infomityServiceCode=MGBOX,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "numericValue")
        if len(str2) > 0:
            service.empty_size = str2[0]

        str1 = "cn=databankContractPCNumber,o=settings,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "numericValue")
        if len(str2) > 0:
            service.contract_pc_num = str2[0]
	
        str1 = "cn=databankSetPCNumber,o=settings,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "numericValue")
        if len(str2) > 0:
            service.set_pc_num = str2[0]
        
        services_list.append(service)
    ticket.services = services_list
    return ticket

def fill_products_list(ticket):
    products_list = []
    PRODUCTS = "o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
    products = exec_ldap_query(LDAPSRV, PRODUCTS, "(objectClass=*)", "productNumber")

    for prod in products:
        SERIAL = "o=devices,productNumber=" + prod + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        serials = exec_ldap_query(LDAPSRV, PRODUCTS, "(objectClass=*)", "serialNumber")
        for serl in serials:
            product = de_classes.Product()
            product.product_number = prod
            product.serial_number = serl

            str1 = "cn=backupEmptySize,serialNumber=" + serl + ",o=devices,productNumber=" + prod + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
            str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "numericValue")
            if len(str2) > 0:
                product.empty_size = str2[0]

            str1 = "cn=backupFullSize,serialNumber=" + serl + ",o=devices,productNumber=" + prod + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
            str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "numericValue")
            if len(str2) > 0:
                product.full_size = str2[0]

            str1 = "cn=useDBRepliIPAddress,serialNumber=" + serl + ",o=devices,productNumber=" + prod + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
            str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "value")
            if len(str2) > 0:
                product.dbrepli_ip = str2[0]

            str1 = "cn=useStorageIPAddress,serialNumber=" + serl + ",o=devices,productNumber=" + prod + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
            str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "value")
            if len(str2) > 0:
                product.storage_ip = str2[0]

            str1 = "cn=backupStopFlag,serialNumber=" + serl + ",o=devices,productNumber=" + prod + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
            str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "enableFlag")
            if len(str2) > 0:
                product.stop_flag = str2[0]
				
            str1 = "serialNumber=" + serl + ",o=devices,productNumber=" + prod + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
            str2 = exec_ldap_query(LDAPSRV, str1, "(objectClass=*)", "createTimestamp")
            if len(str2) > 0:
                product.create_time = str2[0]
				
            products_list.append(product)
    ticket.products = products_list
    return ticket

def rewrite_workbook(workbook, tickets):
	
    wb = copy(workbook)
    worksheet_dbdb = wb.get_sheet(2)
    worksheet_dbdb9a = wb.get_sheet(6)
    worksheet_closed_dbdb = wb.get_sheet(4)
    worksheet_closed_dbdb9a = wb.get_sheet(5)
    fmts = ['@','m/D/YYYY',]

    style_date = xlwt.XFStyle()	
    style_general = xlwt.XFStyle()
    font = xlwt.Font()
    font.bold = False
    style_date.font = font
    style_general.font = font
    style_general.num_format_str = fmts[0]
    style_date.num_format_str = fmts[1]
    borders = xlwt.Borders()
    borders.bottom = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    style_date.borders = borders
    style_general.borders = borders

    rownum_dbdb = 10
    rownum_dbdb9a = 10
    rownum_closed_dbdb = 2
    rownum_closed_dbdb9a = 2
    tickets = sorted(tickets, key=attrgetter('ticket_open_date'), reverse=False)
	
    total_row_dbdb = number_of_good_rows(workbook.sheet_by_name('dbdb'))
    total_row_dbdb9a = number_of_good_rows(workbook.sheet_by_name('dbdb9a'))
    total_row_closed_dbdb = number_of_good_rows(workbook.sheet_by_name('Closed-dbdb'))
    total_row_closed_dbdb9a = number_of_good_rows(workbook.sheet_by_name('Closed-dbdb9a'))
  
    for ticket in tickets:
        if ticket.databank_enabled == True:
            if ticket.sheet_name == "Closed-dbdb9a":
                cur_worksheet = worksheet_closed_dbdb9a
                cur_row = rownum_closed_dbdb9a
                rownum_closed_dbdb9a +=1
            elif ticket.sheet_name == "dbdb9a":
                cur_worksheet = worksheet_dbdb9a
                cur_row = rownum_dbdb9a
                rownum_dbdb9a +=1
            elif ticket.sheet_name == "Closed-dbdb":
                cur_worksheet = worksheet_closed_dbdb
                cur_row = rownum_closed_dbdb
                rownum_closed_dbdb +=1
            elif ticket.sheet_name == "dbdb":
                cur_worksheet = worksheet_dbdb
                cur_row = rownum_dbdb
                rownum_dbdb +=1
		
            if len(ticket.author) <= 1:
			    ticket.author = "Evghenii"
        
            cur_worksheet.write(cur_row, 1, "", style=style_general)
            if type(ticket.ticket_close_date) is datetime.datetime and len(ticket.ticket_close_date.strftime('%m/%d/%Y')) > 1:
                cur_worksheet.write(cur_row, 2, ticket.ticket_close_date.strftime('%m/%d/%Y'), style=style_date)
            else:
                cur_worksheet.write(cur_row, 2, "", style=style_general)
            cur_worksheet.write(cur_row, 3, ticket.ticket_open_date.strftime('%m/%d/%Y'), style=style_date)       
            cur_worksheet.write(cur_row, 4, ticket.author, style=style_general)
            cur_worksheet.write(cur_row, 5, ticket.customer_name, style=style_general)
            cur_worksheet.write(cur_row, 6, ticket.customer_dbid, style=style_general)
            cur_worksheet.write(cur_row, 7, ticket.manual_synchronization, style=style_general)
			
            if type(ticket.ticket_close_date) is datetime.datetime:
                cur_worksheet.write(cur_row, 8, ticket.support_date_end.strftime('%m/%d/%Y'), style=style_date)
            if type(ticket.ticket_close_date) is datetime.datetime:
                cur_worksheet.write(cur_row, 9, ticket.contract_date_start.strftime('%m/%d/%Y'), style=style_date)
            
            cur_worksheet.write(cur_row, 10, ticket.another_db_exist, style=style_general)
            cur_worksheet.write(cur_row, 11, ticket.another_db_status, style=style_general)
            cur_worksheet.write(cur_row, 12, ticket.table_exist_check, style=style_general)
            if not isinstance(ticket.last_db_access_date, basestring):
                cur_worksheet.write(cur_row, 13, ticket.last_db_access_date.strftime('%m/%d/%Y'), style=style_date)    
            if not isinstance(ticket.last_db_sync_date, basestring):
                cur_worksheet.write(cur_row, 14, ticket.last_db_sync_date.strftime('%m/%d/%Y'), style=style_date)
            cur_worksheet.write(cur_row, 15, ticket.log_check, style=style_general)
            cur_worksheet.write(cur_row, 16, ticket.st_server, style=style_general)
            cur_worksheet.write(cur_row, 17, ticket.kim_system_ini_timestamp.strftime('%m/%d/%Y'), style=style_date)
            cur_worksheet.write(cur_row, 18, ticket.image_pilot_version[:8], style=style_general)
            cur_worksheet.write(cur_row, 19, "", style=style_general)
            cur_worksheet.write(cur_row, 20, ticket.comment, style=style_general)
            cur_worksheet.write(cur_row, 21, ticket.action1_db, style=style_general)
            cur_worksheet.write(cur_row, 22, ticket.action1_request, style=style_general)
            cur_worksheet.write(cur_row, 23, ticket.action2_st, style=style_general)
        
            if not isinstance(ticket.action2_db, basestring):
                cur_worksheet.write(cur_row, 24, ticket.action2_db.strftime('%m/%d/%Y'), style=style_date)
        
            cur_worksheet.write(cur_row, 25, ticket.action3_request, style=style_general)
            cur_worksheet.write(cur_row, 26, ticket.call_center_comment, style=style_general)
            
    if rownum_dbdb < total_row_dbdb:
        for rownum in range(rownum_dbdb, total_row_dbdb):
            for colnum in range(1, 27):
                worksheet_dbdb.write(rownum, colnum, "", style=style_general)

    if rownum_dbdb9a < total_row_dbdb9a:
        for rownum in range(rownum_dbdb9a, total_row_dbdb9a):
            for colnum in range(1, 27):
                worksheet_dbdb9a.write(rownum, colnum, "", style=style_general)

    if rownum_closed_dbdb < total_row_closed_dbdb:
        for rownum in range(rownum_closed_dbdb, total_row_closed_dbdb):
            for colnum in range(1, 27):
                worksheet_closed_dbdb.write(rownum, colnum, "", style=style_general)

    if rownum_closed_dbdb9a < total_row_closed_dbdb9a:
        for rownum in range(rownum_closed_dbdb9a, total_row_closed_dbdb9a):
            for colnum in range(1, 27):
                worksheet_closed_dbdb9a.write(rownum, colnum, "", style=style_general)
	
	wb.save(EXCELFILE2)
	os.remove(EXCELFILE)
	os.rename(EXCELFILE2, EXCELFILE)

def get_client_version(server, tickets):
    ticket_list = []
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    cmd = "sh /infomity-mainte/bin/stcheck/KimSystemini_chk.sh"
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    for ticket in tickets:
        for x in range(len(teststr)):
            str1 = re.split(",", teststr[x])
            if str1[0] == ticket.customer_dbid[:7] and str1[1] == ticket.customer_dbid[7:]:   
                if len(ticket.image_pilot_version) < 1:
                    ticket.image_pilot_version = str1[3]
                if isinstance(ticket.kim_system_ini_timestamp, basestring):
                    if len(ticket.kim_system_ini_timestamp) < 1:
                        ticket.kim_system_ini_timestamp = str_to_date1(str1[2][:10])
                elif type(ticket.kim_system_ini_timestamp) is datetime.datetime:
                    if len(ticket.kim_system_ini_timestamp.strftime("%B,%d,%Y")) < 1:
                        ticket.kim_system_ini_timestamp = str_to_date1(str1[2][:10])

        ticket_list.append(ticket)
    ssh.close()
    return ticket_list
		
def get_db_repli(server, ticket):
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    cmd="/usr/local/pgsql/DBRepli/DBRepliDelay.sh " + ticket.customer_dbid
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    ticket.manual_synchronization = "NG"
    if len(teststr) == 1 and len(teststr[0]) > 2:
        ticket.tmp = teststr[0]
        if teststr[0] == "DB DOWN":
            ticket.manual_synchronization = "NG"
        else:
            str1 = re.split(",", teststr[0])
            str2 = re.split(":", teststr[0])
            if int(str1[1]) == 0 and int(str2[0]) < 100:
                ticket.manual_synchronization = "OK"
    ssh.close()
    return ticket
		
def get_another_database(server, ticket):
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    cmd="/usr/local/pgsql/bin/psql -U postgres -t -p 5433 postgres -c \"SELECT datname FROM pg_database where datname like '"+ticket.customer_dbid[:7]+"%';\" "
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    if len(teststr) > 2:
        for x in range(1, len(teststr)-1):
            ticket.another_db_exist = ticket.another_db_exist + " " + teststr[x][:-1]
        ticket.another_db_status = "NG"
    else:
        ticket.another_db_exist = ""
	"""
	#checking customer database status
	cmd = "/usr/local/pgsql/bin/psql -U postgres -p 5433 " + ticket.customer_dbid[:7] + " -t -c \"select count(*) from kim.yuubin_tbl;\""
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    #if len(teststr) > 0:
	
    cmd = "/usr/local/pgsql/bin/psql -U postgres -p 5433 " + ticket.customer_dbid[:7]+" -t -c \"select count(*) from _c1.sl_status;\""
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    #if len(teststr) > 0:
	"""

    #if True:
    cmd = "/usr/local/pgsql/bin/psql -U postgres -p 5433 " + ticket.customer_dbid + " -t -c \"select st_last_event_ts from _c1.sl_status ;\""
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    if len(teststr) > 0:
        ticket.last_db_access_date = str_to_date1(teststr[0][1:11])
	
    cmd = "/usr/local/pgsql/bin/psql -U postgres -p 5433 " + ticket.customer_dbid + " -t -c \"select st_last_received_event_ts from _c1.sl_status ;\""
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    if len(teststr) > 0:
        ticket.last_db_sync_date = str_to_date1(teststr[0][1:11])
    ssh.close()	
    return ticket

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
            str = i.pretty_print2()
            if str is not None and len(str) > 0:
                values.append(str)
    except ldap.LDAPError, e:
        pass	
    return values	

def comment_problem(ticket):
    comment = ""
    if len(ticket.services) > 0:
        for serv in ticket.services:		
            if serv.service_name == "DATABANK":
                ticket.support_date_end = str_to_date1(serv.service_end[:8])
                ticket.contract_date_start = str_to_date1(serv.contract_start[:8])
              
                if ticket.support_date_end < datetime.datetime.today():
                    ticket.comment = ticket.comment + " Support date expired " + ticket.support_date_end.strftime("%m/%d/%Y")
                if ticket.contract_date_start < datetime.datetime.today():
                    ticket.comment = ticket.comment + " Contract date expired " + ticket.contract_date_start.strftime("%m/%d/%Y")
    
    if len(ticket.products) > 0:
        for prod in ticket.products:
            siteid = ticket.customer_dbid[:7] + prod.product_number + prod.serial_number
            
            if ticket.customer_dbid == siteid and prod.empty_size == "0":
                print siteid + "  " + prod.empty_size
                ticket.comment = ticket.comment + " Customer has not free disk space."	
	return ticket			
	
def fill_new_tickets(ticket):
    #fill Customer name
    CUSTOMER_NAME_QUERY="nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
    str = exec_ldap_query(LDAPSRV, CUSTOMER_NAME_QUERY, "(objectClass=*)", "medicalInstitutionName")
    if len(str) > 0:
        ticket.customer_name = str[0]

	#Fill cussomer phone number	
    TEL_NUMBER_QUERY="nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
    str = exec_ldap_query(LDAPSRV, TEL_NUMBER_QUERY,"(objectClass=*)", "telephoneNumber")
    if len(str) > 0:
        ticket.phone_number = str[0]
    
    PACKAGE_INFO_QUERY = "o=packageContracts,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
    str = exec_ldap_query(LDAPSRV, PACKAGE_INFO_QUERY,"contractFlag=TRUE", "infomityServicePackageCode")
    if len(str) > 0:
        ticket.package_info = str[0]
        
    #check if it is new ticket
    ifNoCreateionDate = False
    if isinstance(ticket.ticket_open_date, basestring):
        if len(ticket.ticket_open_date) < 1:
            ifNoCreateionDate = True         
    elif type(ticket.ticket_open_date) is datetime.datetime:
        if len(ticket.ticket_open_date.strftime("%B,%d,%Y")) < 1:
            ifNoCreateionDate = True
		
    if ifNoCreateionDate is True:
        ticket.ticket_open_date = datetime.datetime.today()
        ticket = fill_services_list(ticket)
        ticket = fill_products_list(ticket) 
	
	#Check second dababase	
    if ticket.image_pilot_version[:3] == "1.7":
        ticket = get_another_database(DBDB9ASRV, ticket)
    else:
        ticket = get_another_database(DBDBSRV, ticket)

    ticket.databank_enabled = False		
    today = datetime.datetime.today()
    for serv in ticket.services:
        if serv.service_name == "DATABANK":
            ticket.databank_enabled = True
            ticket.support_date_end = serv.service_end
            ticket.contract_date_start = serv.contract_start
            if serv.contract_end < today:
                ticket.contract_expired = serv.contract_end

    if ticket.image_pilot_version[:3] == "1.7":
        ticket = get_db_repli(DBDB9ASRV, ticket)
    else:
        ticket = get_db_repli(DBDBSRV, ticket)

    return ticket
	
def fill_and_write_new_tickets(workbook, tickets):
    tickets = get_client_version(DBSTSRV, tickets)
    for ticket in tickets:
        ifCreateionDate = False
        if isinstance(ticket.ticket_open_date, basestring):
            if len(ticket.ticket_open_date) <= 1:
                ifCreateionDate = True         
        elif type(ticket.ticket_open_date) is datetime.datetime:
            if len(ticket.ticket_open_date.strftime("%B,%d,%Y")) <= 1:
                ifCreateionDate = True
        
        if ifCreateionDate is True:
            ticket = fill_new_tickets(ticket)
            ticket = comment_problem(ticket)

    rewrite_workbook(workbook, tickets)

def check_if_replicated(server, tickets):
    ticket_list = []    
    today = datetime.datetime.today()
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    cmd = "sh /infomity-mainte/bin/stcheck/KimSystemini_chk.sh | grep -v \"KimSystem.ini\""
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    for x in range(len(teststr)):
        str1 = re.split(",", teststr[x])
        str2 = string.split(str1[2])
        kmidate = str_to_date1(str2[0])
        for ticket in tickets:
            if str1[0] == ticket.customer_dbid[:7] and str1[1] == ticket.customer_dbid[7:]:   
                today = datetime.datetime.now()
                today = str_to_date1(str(today.year)+"/"+str(today.month)+"/"+str(today.day))
                diff_days = abs((today - kmidate).days)
                if diff_days < 4:
                    ticket.customer_dbid = str1[0] + str1[1]
					#cheking if DB well replicated
                    if ticket.image_pilot_version[:3] == "1.7":
                        ticket = get_db_repli(DBDB9ASRV, ticket)
                    else:
                        ticket = get_db_repli(DBDBSRV, ticket)
                    if ticket.manual_synchronization == "OK":
                        ticket.ticket_close_date = datetime.datetime.today()
                        if server == "10.227.9.223":
                            ticket.sheet_name = "Closed-dbdb9a"
                        else:
                            ticket.sheet_name = "Closed-dbdb"
                        ticket_list.append(ticket)

    ssh.close()
    return ticket_list

def print_date(function_name):
    d = datetime.datetime.today()
    print str(d) + " " + function_name
	
def print_list(tickets):
    for ticket in tickets:
        print ticket.customer_dbid
	
#read excel file
print_date('begin')

xl_workbook = xlrd.open_workbook(EXCELFILE, formatting_info=True)	

closed_dbdblist = de_read_sheet(xl_workbook, 'Closed-dbdb', 3)
print_date('de_utils.de_utils.de_read_sheet')

closed_dbdb9alist = de_read_sheet(xl_workbook,'Closed-dbdb9a', 3)
print_date('de_utils.de_utils.de_read_sheet')

dbdblist = de_read_sheet(xl_workbook,'dbdb', 11)
print_date('de_utils.de_read_sheet')

dbdb9alist = de_read_sheet(xl_workbook,'dbdb9a', 11)
print_date('de_utils.de_read_sheet')


#close replicated tickets
dbdblist_replicated = check_if_replicated(DBSTSRV, dbdblist)
print_date('check_if_replicated')

dbdb9alist_replicated = check_if_replicated(DBSTSRV, dbdb9alist)
print_date('check_if_replicated')

dbdblist = remove_dublicates(dbdblist, dbdblist_replicated)
print_date('remove_dublicates')

dbdb9alist = remove_dublicates(dbdb9alist, dbdb9alist_replicated)
print_date('remove_dublicates')

closed_dbdblist = assembly_two_list(closed_dbdblist, dbdblist_replicated)
print_date('assembly_two_list')

closed_dbdb9alist = assembly_two_list(closed_dbdb9alist, dbdb9alist_replicated)
print_date('assembly_two_list')

#check new tickets on servers
dbdb_new = get_new_tickest_from_db(DBDBSRV, DBDB_FILTER)
print_date('get_new_tickest_from_db')

dbdb9a_new = get_new_tickest_from_db(DBDB9ASRV, DBDB9A_FILTER)
print_date('get_new_tickest_from_db')

dbst_new = get_new_tickest_from_st(DBSTSRV, DBST_FILTER)
print_date('get_new_tickest_from_st')

dbdb_new = remove_dublicates(dbdb_new, dbdblist)
print_date('remove_dublicates')

dbst_new = remove_dublicates(dbst_new, dbdblist)
dbst_new = remove_dublicates(dbst_new, dbdb_new)

dbdb9a_new = remove_dublicates(dbdb9a_new, dbdb9alist)
print_date('remove_dublicates')

dbst_new = remove_dublicates(dbst_new, dbdb9alist)
print_date('remove_dublicates')

dbst_new = remove_dublicates(dbst_new, dbdb9a_new)
print_date('remove_dublicates')

common_list = assembly_two_list(dbdblist, dbdb_new)
common_list = assembly_two_list(common_list, dbdb9alist)
common_list = assembly_two_list(common_list, dbdb9a_new)
common_list = assembly_two_list(common_list, dbst_new)
common_list = assembly_two_list(common_list, closed_dbdblist)
common_list = assembly_two_list(common_list, closed_dbdb9alist)

fill_and_write_new_tickets(xl_workbook, common_list)
print_date('fill_and_write_new_tickets')
