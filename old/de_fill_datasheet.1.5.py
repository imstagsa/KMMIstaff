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
DBST_FILTER='app\\dbdst_filter.txt'
DBDB_FILTER='app\\dbdb_filter.txt'
DBDB9A_FILTER='app\\dbdb9a_filter.txt'

def get_new_tickest_from_st(server, filter_file):
    ticket_list = []
    today = datetime.datetime.today()
    exclude_list = read_filter_from_file(filter_file)
    cmd = 'sh /infomity-mainte/bin/stcheck/KimSystemini_chk.sh | egrep -iv \"'+exclude_list+'\" '
    teststr = exec_command_via_ssh(server, cmd)	
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
                if str3[:3] == "1.7" or str3[:3] == "1.8":
                    ticket.sheet_name = "dbdb9a"
                else:
                    ticket.sheet_name = "dbdb"
                ticket_list.append(ticket)
    return ticket_list
		
def get_new_tickest_from_db(server, filter_file):
    ticket_list = []
    today = datetime.datetime.today() 
    exclude_list = read_filter_from_file(filter_file)
    cmd = 'cd /usr/local/pgsql/data/infomity-mainte/log/; cat `ls -tr infomity-userDBRepliDelayPingCSV-ForStatistics-' + str(today.year) + '* | tail -n 1` | egrep -iv \"'+exclude_list+'\" | grep ",[0-1]"'
    teststr = exec_command_via_ssh(server, cmd)	
    for x in range(len(teststr)):
        str1 = re.split("\|", teststr[x])
        if len(str1) == 3:
            str2 = re.split(":", str1[2])
            if int(str2[0]) > 100:
                ticket = de_classes.Ticket()
                ticket.customer_dbid = str1[1]
                if server == '10.227.9.223':
                    ticket.sheet_name = 'dbdb9a'
                else:
                    ticket.sheet_name = 'dbdb'
                #Checking if DB exist
                cmd = 'grep ' + ticket.customer_dbid +  '  /mount-dir/pgsql/data/pg_hba.conf | wc -l'
                teststr1 = exec_command_via_ssh(server, cmd)
                if int(teststr1[0]) == 0:
                    ticket.comment = ticket.comment + 'DB doesnt created. Custormer doesnt configured. Please check. '
                else:
                    cmd = 'grep ' + ticket.customer_dbid +  ' /mount-dir/pgsql/data/pg_hba.conf'
                    teststr1 = exec_command_via_ssh(server, cmd)
                    if teststr1[0][:1] == '#':
                         ticket.comment = ticket.comment + 'Customer already commented from pg_hba.conf. '
                ticket_list.append(ticket)

				
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
'''
def get_dublicates(list1, list2):
    listidx = []
    list_tmp = []
    for x in range(len(list1)):
        for y in range(len(list2)):
            if list1[x].customer_dbid == list2[y].customer_dbid and list1[x].ticket_open_date == list2[y].ticket_open_date:
                list_tmp.append(list1[x])

    return list_tmp	
'''	
	
#return one ticket list with elements from firsts and second lists
def assembly_two_lists(list1, list2):
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
    worksheet_dbdb9a = wb.get_sheet(5)
    worksheet_closed_dbdb = wb.get_sheet(3)
    worksheet_closed_dbdb9a = wb.get_sheet(4)
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
                #print ticket.customer_dbid
                cur_worksheet = worksheet_closed_dbdb9a
                cur_row = rownum_closed_dbdb9a
                rownum_closed_dbdb9a +=1
            elif ticket.sheet_name == "dbdb9a":
                cur_worksheet = worksheet_dbdb9a
                cur_row = rownum_dbdb9a
                rownum_dbdb9a +=1
                #print "dbdb9a:"+ticket.customer_dbid
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

            if type(ticket.support_date_end) is datetime.datetime:
                cur_worksheet.write(cur_row, 8, ticket.support_date_end.strftime('%m/%d/%Y'), style=style_date)
            if type(ticket.contract_date_start) is datetime.datetime:
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
            if not isinstance(ticket.kim_system_ini_timestamp, basestring):
                cur_worksheet.write(cur_row, 17, ticket.kim_system_ini_timestamp.strftime('%m/%d/%Y'), style=style_date)
            cur_worksheet.write(cur_row, 18, ticket.image_pilot_version[:8], style=style_general)
            cur_worksheet.write(cur_row, 19, "", style=style_general)
            cur_worksheet.write(cur_row, 20, ticket.comment, style=style_general)
            cur_worksheet.write(cur_row, 21, ticket.action1_db, style=style_general)
            cur_worksheet.write(cur_row, 22, ticket.action1_request, style=style_general)
            if not isinstance(ticket.comment_out_from_st, basestring):
                cur_worksheet.write(cur_row, 23, ticket.comment_out_from_st.strftime('%m/%d/%Y'), style=style_date)
            if not isinstance(ticket.comment_out_from_pg_hba, basestring):
                cur_worksheet.write(cur_row, 24, ticket.comment_out_from_pg_hba.strftime('%m/%d/%Y'), style=style_date)
        
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
    try:
        wb.save(EXCELFILE2)
    except ValueError,e:
        print e
	
    os.remove(EXCELFILE)
    os.rename(EXCELFILE2, EXCELFILE)

def get_client_version(server, tickets):
    ticket_list = []
    cmd = "sh /infomity-mainte/bin/stcheck/KimSystemini_chk.sh"
    teststr = exec_command_via_ssh(server, cmd)	
    for ticket in tickets:
        for x in range(len(teststr)):
            str1 = re.split(",", teststr[x])
            if str1[0] == ticket.customer_dbid[:7] and str1[1] == ticket.customer_dbid[7:]:   
                #print len(str1)
                if len(ticket.image_pilot_version) < 1:
                    if len(str1) > 3: 
                        ticket.image_pilot_version = str1[3]
                if isinstance(ticket.kim_system_ini_timestamp, basestring):
                    if len(ticket.kim_system_ini_timestamp) < 1:
                        ticket.kim_system_ini_timestamp = str_to_date1(str1[2][:10])
                elif type(ticket.kim_system_ini_timestamp) is datetime.datetime:
                    if len(ticket.kim_system_ini_timestamp.strftime("%B,%d,%Y")) < 1:
                        ticket.kim_system_ini_timestamp = str_to_date1(str1[2][:10])
        ticket_list.append(ticket)
    return ticket_list
		
def get_db_repli(server, ticket):
    cmd="/usr/local/pgsql/DBRepli/DBRepliDelay.sh " + ticket.customer_dbid
    teststr = exec_command_via_ssh(server, cmd)	
    ticket.manual_synchronization = "NG"
    if len(teststr) == 1 and len(teststr[0]) > 2:
        #ticket.tmp = teststr[0]
        if teststr[0] == "DB DOWN":
            ticket.manual_synchronization = "NG"
            ticket.comment = ticket.comment + " Database down. "
        else:
            str1 = re.split(",", teststr[0])
            str2 = re.split(":", teststr[0])
            ticket.last_db_sync_hours = int(str2[0])
            #if int(str1[1]) == 0 and int(str2[0]) < 100:
            if int(str2[0]) < 100:
                ticket.manual_synchronization = "OK"
    return ticket
		
def get_another_database(server, ticket):
    cmd="/usr/local/pgsql/bin/psql -U postgres -t -p 5433 postgres -c \"SELECT datname FROM pg_database where datname like '"+ticket.customer_dbid[:7]+"%';\" "
    teststr = exec_command_via_ssh(server, cmd)	
    if len(teststr) > 2:
        for x in range(0, len(teststr)-1):
            if " "+ticket.customer_dbid != teststr[x][:-1]:
                ticket.another_db_exist = ticket.another_db_exist + teststr[x][:-1]
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
    teststr = exec_command_via_ssh(server, cmd)
    if len(teststr) > 0:
        ticket.last_db_access_date = str_to_date1(teststr[0][1:11])
	
    cmd = "/usr/local/pgsql/bin/psql -U postgres -p 5433 " + ticket.customer_dbid + " -t -c \"select st_last_received_event_ts from _c1.sl_status ;\""
    teststr = exec_command_via_ssh(server, cmd)
    if len(teststr) > 0:
        ticket.last_db_sync_date = str_to_date1(teststr[0][1:11])
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
                    ticket.comment = ticket.comment + " Support date expired " + ticket.support_date_end.strftime("%m/%d/%Y") + ". "
                if type(ticket.contract_expired) is datetime.datetime and len(ticket.contract_expired.strftime('%m/%d/%Y')) > 1:
                    if ticket.contract_expired < datetime.datetime.today():
                        ticket.comment = ticket.comment + " Contract date expired " + ticket.contract_expired.strftime("%m/%d/%Y")+ ". "
    
    if len(ticket.products) > 0:
        for prod in ticket.products:
            siteid = ticket.customer_dbid[:7] + prod.product_number + prod.serial_number
            
            if ticket.customer_dbid == siteid and prod.empty_size == "0":
                ticket.comment = ticket.comment + "Customer does not has free disk space. "

    versions = ['1.5', '1.6', '1.7', '1.8']
    dbdb9a_versions = ['1.7', '1.8']
    if ticket.image_pilot_version[:3] in versions:
        if ticket.sheet_name == "dbdb" and ticket.image_pilot_version[:3] in dbdb9a_versions:
            ticket.comment = ticket.comment + "Incorrect ImagePilot Version. "
        elif ticket.sheet_name == "dbdb9a" and ticket.image_pilot_version[:3] not in  dbdb9a_versions :
            ticket.comment = ticket.comment + "Incorrect ImagePilot Version. "
    else:
        ticket.comment = ticket.comment + "Incorrect ImagePilot Version. "
    
    if ticket.last_db_sync_hours > 100:
        ticket.comment = ticket.comment + "Request TCC to check condition of ImagePilot PC and reboot if needed, since there is no access/replication from facility to ST/DB server. "
	
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
    if ticket.image_pilot_version[:3] == "1.7" or ticket.image_pilot_version[:3] == "1.8":
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

    if ticket.image_pilot_version[:3] == "1.7" or ticket.image_pilot_version[:3] == "1.8":
        ticket = get_db_repli(DBDB9ASRV, ticket)
    elif ticket.image_pilot_version[:3] == "1.6" or ticket.image_pilot_version[:3] == "1.5":
        ticket = get_db_repli(DBDBSRV, ticket)
    else:
        ticket = get_db_repli(DBDBSRV, ticket)
        ticket = get_db_repli(DBDB9ASRV, ticket)

    return ticket
	
def fill_and_write_new_tickets(workbook, tickets):
    print_date("get_client_version") 
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
    print_date("rewrite_workbook") 
    rewrite_workbook(workbook, tickets)

def check_if_commented_out(list1, list2, sheet_name):
    list_tmp = []
    for x in range(len(list1)):
        for y in range(len(list2)):
            if list1[x].customer_dbid == list2[y].customer_dbid:
                list1[x].comment = list1[x].comment + "Commented out."
                list1[x].ticket_close_date = datetime.datetime.today()
                list1[x].sheet_name = sheet_name
                list_tmp.append(list1[x])

    return list_tmp
	
def check_if_replicated(server, tickets):
    ticket_list = []    
    today = datetime.datetime.today()
    cmd = "sh /infomity-mainte/bin/stcheck/KimSystemini_chk.sh | grep -v \"KimSystem.ini\""
    teststr = exec_command_via_ssh(server, cmd)
    for x in range(len(teststr)):
        str1 = re.split(",", teststr[x])
        str2 = string.split(str1[2])
        kmidate = str_to_date1(str2[0])
        #print teststr[x]
        for ticket in tickets:
            if str1[0] == ticket.customer_dbid[:7] and str1[1] == ticket.customer_dbid[7:]:   
                today = datetime.datetime.now()
                today = str_to_date1(str(today.year)+"/"+str(today.month)+"/"+str(today.day))
                diff_days = abs((today - kmidate).days)
                if diff_days < 4:
                    ticket.customer_dbid = str1[0] + str1[1]
					#cheking if DB well replicated
                    if ticket.image_pilot_version[:3] == "1.7" or ticket.image_pilot_version[:3] == "1.8":
                        ticket = get_db_repli(DBDB9ASRV, ticket)
                    else:
                        ticket = get_db_repli(DBDBSRV, ticket)
                    if ticket.manual_synchronization == "OK" and ticket.last_db_sync_hours < 100:
                        ticket.ticket_close_date = datetime.datetime.today()
                        if ticket.image_pilot_version[:3] == "1.7" or ticket.image_pilot_version[:3] == "1.8":
                            ticket.sheet_name = "Closed-dbdb9a"
                        else:
                            ticket.sheet_name = "Closed-dbdb"
                        ticket_list.append(ticket)
    return ticket_list

def comment_out_from_pg_hba(server, ticket):
    file_name='/usr/local/pgsql/data/pg_hba.conf'
    today = datetime.datetime.today()
    file_name_bkp=file_name+"_"+today.strftime('%m%d%Y%M%S%f')[:-3]
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    cmd = 'cp '+file_name+' '+file_name_bkp
    stdin, stdout, stderr = ssh.exec_command(cmd)
    cmd = 'sed -i  \'s/^hostssl\ '+ticket.customer_dbid+'/#hostssl\ '+ticket.customer_dbid+'/\' /usr/local/pgsql/data/pg_hba.conf'
    stdin, stdout, stderr = ssh.exec_command(cmd)
    cmd = 'runuser -l postgres -c \'pg_ctl reload\''
    stdin, stdout, stderr = ssh.exec_command(cmd)
    ssh.close()

def add_to_filter_in_dbst(server, ticket):
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    #cmd = 'str=`cat /infomity-mainte/bin/stcheck/SiteFilter.txt`; echo $str"|'+ticket.customer_dbid[:7]+','+ticket.customer_dbid[7:]+'" > /infomity-mainte/bin/stcheck/SiteFilter.txt'
    #stdin, stdout, stderr = ssh.exec_command(cmd)
    cmd = 'echo "'+ticket.customer_dbid[:7]+','+ticket.customer_dbid[7:]+'" >> /Admin/scripts/site_filter.txt'
    stdin, stdout, stderr = ssh.exec_command(cmd)
    ssh.close()

def add_to_local_filter_file(filter_file, ticket):
    filter_file2 = filter_file + '2'
    f1 = open(filter_file, 'r')
    f2 = open(filter_file2, 'w')
    line = f1.readline()
    line = line + '|'+ticket.customer_dbid[:7]+","+ticket.customer_dbid[7:]
    f2.write(line)
    f1.close()
    f2.close()
    os.remove(filter_file)
    os.rename(filter_file2, filter_file)
'''	
def comment_out(tickets):
    get_client_version(DBSTSRV, tickets)
    for ticket in tickets:
        print 'Found customer ' + ticket.customer_dbid + ' for comment.'
        ticket.ticket_close_date = datetime.datetime.today()
        ticket.comment_out_from_pg_hba = datetime.datetime.today()
        ticket.comment = ticket.comment + ' Commented out from pg_hba.conf.'
        if ticket.image_pilot_version[:3] == "1.7":
            comment_out_from_pg_hba(DBDB9ASRV, ticket)
            add_to_local_filter_file(DBDB9A_FILTER, ticket)
        else:
            comment_out_from_pg_hba(DBDBSRV, ticket)
            add_to_local_filter_file(DBDB_FILTER, ticket)
            add_to_local_filter_file(DBST_FILTER, ticket)			
        add_to_filter_in_dbst(DBSTSRV, ticket)
        
	#return tickets
	
def get_ids_for_comment_out(tickets, sheet_name):
    ticket_list = []
    ldapstr1 = 'Contract expired LDAP and Image Pilot disconnected'
    ldapstr2 = 'Contract expired turned off in LDAP'
    ldapstr3 = 'Contract expired, LDAP and Image Pilot disconnected'
    ldapstr4 = 'off on the PC and in LDAP.'
    for ticket in tickets:
        if (ldapstr1 in ticket.call_center_comment) or (ldapstr2 in ticket.call_center_comment) or (ldapstr3 in ticket.call_center_comment) or (ldapstr4 in ticket.call_center_comment):
            print 'Fount ID for comment: ' + ticket.customer_dbid
            ticket.sheet_name = sheet_name
            ticket.ticket_close_date = datetime.datetime.today()
            ticket_list.append(ticket)
    return ticket_list
'''
def rewrite_commented_tab(file, tickets):
    workbook = xlrd.open_workbook(file, formatting_info=True)
    wb = copy(workbook)
    worksheet_comment = wb.get_sheet(6)
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

    rownum_comment = 2
    #tickets = sorted(tickets, key=attrgetter('ticket_open_date'), reverse=False)
    total_row_comment = number_of_good_rows(workbook.sheet_by_name('Commented-out'))
  
    for ticket in tickets:
        cur_row = rownum_comment
        rownum_comment +=1
        
        if type(ticket.ticket_close_date) is datetime.datetime and len(ticket.ticket_close_date.strftime('%m/%d/%Y')) > 1:
            worksheet_comment.write(cur_row, 2, ticket.ticket_close_date.strftime('%m/%d/%Y'), style=style_date)
        else:
            worksheet_comment.write(cur_row, 2, "", style=style_general)
		
        if not isinstance(ticket.ticket_open_date, basestring):
		    worksheet_comment.write(cur_row, 3, ticket.ticket_open_date.strftime('%m/%d/%Y'), style=style_date)
        worksheet_comment.write(cur_row, 4, ticket.author, style=style_general)
        worksheet_comment.write(cur_row, 5, ticket.customer_name, style=style_general)
        worksheet_comment.write(cur_row, 6, ticket.customer_dbid, style=style_general)
        worksheet_comment.write(cur_row, 7, ticket.image_pilot_version[:8], style=style_general)
        worksheet_comment.write(cur_row, 8, ticket.comment, style=style_general)  
            
    if rownum_comment < total_row_comment:
        for rownum in range(rownum_comment, total_row_comment):
            for colnum in range(1, 8):
                worksheet_comment.write(rownum, colnum, "", style=style_general)

    try:
        wb.save(EXCELFILE2)
    except ValueError,e:
        print e
	
    os.remove(file)
    os.rename(EXCELFILE2, file)
	
def comment_out_from_tab(tickets):
    for ticket in tickets:
        if type(ticket.ticket_close_date) is not datetime.datetime:
            get_client_version(DBSTSRV, tickets)
            print "Found customer " + ticket.customer_dbid + " for comment."
            if ticket.image_pilot_version[:3] == "1.7" or ticket.image_pilot_version[:3] == "1.8":
                comment_out_from_pg_hba(DBDB9ASRV, ticket)
                add_to_local_filter_file(DBDB9A_FILTER, ticket)
            else:
                comment_out_from_pg_hba(DBDBSRV, ticket)
                add_to_local_filter_file(DBDB_FILTER, ticket)
            add_to_local_filter_file(DBST_FILTER, ticket)
            add_to_filter_in_dbst(DBSTSRV, ticket)
            ticket.ticket_close_date = datetime.datetime.today()
		
	#return tickets


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

print_date('begin')
try:

    #mount network drive
    mount_x()
	
    #backing up excel file
    copy_file(EXCELFILE, EXCEFILELBKP)
	
    #read excel file
    xl_workbook = xlrd.open_workbook(EXCELFILE, formatting_info=True)	
    
	#Comment out customers
    commentlist = de_read_sheet_comment_out(xl_workbook,"Commented-out", 2)
    print_date('de_read_sheet_comment_out')
    comment_out_from_tab(commentlist)
    print_date('comment_out_from_tab')    
    rewrite_commented_tab(EXCELFILE, commentlist)
    print_date('rewrite_commented_tab')
	
    closed_dbdblist = de_read_sheet(xl_workbook, 'Closed-dbdb', 2)
    print_date('de_read_sheet')
	  
    closed_dbdb9alist = de_read_sheet(xl_workbook,'Closed-dbdb9a', 2)
    print_date('de_read_sheet')

    dbdblist = de_read_sheet(xl_workbook,'dbdb', 10)
    print_date('de_read_sheet')
	
    dbdb9alist = de_read_sheet(xl_workbook,'dbdb9a', 10)
    print_date('de_read_sheet')
    
    #close commented out customers
    dbdblist_commented_out = check_if_commented_out(dbdblist, commentlist, "Closed-dbdb")
    #print_list(dbdblist_commented_out)
    dbdblist = remove_dublicates(dbdblist, dbdblist_commented_out)
    closed_dbdblist = assembly_two_lists(closed_dbdblist, dbdblist_commented_out)
    #print_list(closed_dbdblist)
    print_date('check_if_commented_out')
	
    #close commented out customers
    dbdb9alist_commented_out = check_if_commented_out(dbdb9alist, commentlist, "Closed-dbdb9a")
    dbdb9alist = remove_dublicates(dbdb9alist, dbdb9alist_commented_out)
    closed_dbdb9alist = assembly_two_lists(closed_dbdb9alist, dbdb9alist_commented_out)
    print_date('check_if_commented_out')
    
    #close replicated tickets
    print_date("Starting check_if_replicated.")
    dbdblist_replicated = check_if_replicated(DBSTSRV, dbdblist)
    print_date('check_if_replicated has been finished')

    dbdb9alist_replicated = check_if_replicated(DBSTSRV, dbdb9alist)
    print_date('check_if_replicated')

    dbdblist = remove_dublicates(dbdblist, dbdblist_replicated)
    print_date('remove_dublicates')

    dbdb9alist = remove_dublicates(dbdb9alist, dbdb9alist_replicated)
    print_date('remove_dublicates')

    closed_dbdblist = assembly_two_lists(closed_dbdblist, dbdblist_replicated)
    print_date('assembly_two_lists')

    closed_dbdb9alist = assembly_two_lists(closed_dbdb9alist, dbdb9alist_replicated)
    print_date('assembly_two_lists')
	
	#get ids for comment out from db
    #dbdb_commented = get_ids_for_comment_out(dbdblist, 'Closed-dbdb')
    #dbdblist = remove_dublicates(dbdblist, dbdb_commented)
    #comment_out(dbdb_commented)
    #closed_dbdblist = assembly_two_lists(closed_dbdblist, dbdb_commented)
    
    #dbdb9a_commented = get_ids_for_comment_out(dbdb9alist, 'Closed-dbdb9a')
    #dbdb9alist = remove_dublicates(dbdb9alist, dbdb9a_commented)
    #comment_out(dbdb9a_commented)
    #closed_dbdb9alist = assembly_two_lists(closed_dbdb9alist, dbdb9a_commented)
    
    #check new tickets on servers
    dbdb_new = get_new_tickest_from_db(DBDBSRV, DBDB_FILTER)
    print_list(dbdb_new)
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

    common_list = assembly_two_lists(dbdblist, dbdb_new)
    common_list = assembly_two_lists(common_list, dbdb9alist)
    common_list = assembly_two_lists(common_list, dbdb9a_new)
    common_list = assembly_two_lists(common_list, dbst_new)
    common_list = assembly_two_lists(common_list, closed_dbdblist)
    common_list = assembly_two_lists(common_list, closed_dbdb9alist)

    fill_and_write_new_tickets(xl_workbook, common_list)
    print_date('fill_and_write_new_tickets')
    
    unmount_x()
	
    send_email(SMTPSERVER, SENDFROM, SENDTO, "KMMI job status OK", "Execution C:\\Python27\\app\\de_fill_datasheet.py: OK")

except ValueError, e:
    print e

