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
from de_utils import str_to_date3
from de_utils import str_to_date4
from de_utils import exec_ldap_query
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from xlrd import open_workbook
from xlutils.copy import copy
from operator import attrgetter
from operator import itemgetter

DBSTSRV = '10.227.9.205'
DBDBSRV = '10.227.9.219'
DBDB9ASRV = '10.227.9.223'
LDAPSRV = "ldap://10.227.9.233"
DBDBSUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb.sh'
DBDB9ASUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb9a.sh'
EXCELFILE = 'app\\04-03_Replication_failure_control_doc.xls'
EXCELFILE2 = 'app\\04-03_Replication_failure_control_doc2.xls'
DBDB_EXCLUDE = 'egrep -iv \"2000010|2008977|2004785|2008195|2009247|2009275|2009612|2008203|2008261|2008470|delay|DEMO|DBRepli\"'
DBDB9A_EXCLUDE = 'egrep -iv \"2000010|2008977|2004785|2008195|2009247|2009275|2009612|2008203|2008261|2012410|2012412|2012413|2012414|2012415|2012517|delay|DEMO|DBRepli\"'
DBST_EXCLUDE = 'egrep -iv \"backup|1111111|2000010|2008977|2004785|2008195|2009247|2009275|2009612|2008203|2008261|2008470|DEMO|2012410|2012412|2012413|2012414|2012415|2012517|_2009451_backup20130926|H1USTEST|KMMG|mhustest|KimSystem.ini|backup_|2012102|2012409|2012411|2005040,59023265040A10102,|2005345,5902300617016024,|2007524,5902320573173975,|2007659,590232001139,|2007885,59023186645916137,|2008565,590232300001,|2008609,5902336241924962,|2008770,5902326130165507,|2008779,5902326519102819,|2008816,5902330176455434,|2008834,5902326308408561,|2008880,5902330594293755,|2008887,5902330800635855,|2008927,5902326801068366,|2008955,5902326235214421,|2008968,5902326906255933,|2008972,5902330168393170,|2008973,5902330628987942,|2009050,5902326247380041,|2009101,5902326281050326,|2009102,5902326515329586,|2009124,5902326323067626,|2009125,5902326644898398,|2009126,5902326965607020,|2009133,590236981485420,|2009162,5902326390764034,|2009197,5902326550614006,|2009439,5902330996320817,|2009451,5902330833817731,|2009451,5902334110896529,|2009676,59023300001,|2009676,5902330479449308,|2009749,5902330691271287,|2009790,5902330504560097,|2009790,5902330930792648,|2010077,590233456509572,|2010222,5902334153244039,|2010516,5902319108040445,|2010529,5902336713833534,|2010830,590233400001,|2009288,5902330351771896,|2009366,5902330808628696,|2008113,59020131680,|2008107,5902320PL12UA04819WW,|2001680,590230099C0676,|2009070,5902326489260566,|2007960,590231901455,|2002755,5902300MJ00717,|2009543,5902323956247757,|2009642,5902330105810346,|2008031,5902319374978909,|2008971,5902330358998116,|2009518,5902330860249699,|2008667,5902319783501179,|2008946,5902330707669365,|2008840,5902326494911924,|2009576,59023262UA1290GV0,|2008189,590231901732\"'

def get_new_tickest_from_st(server):
    i = 0
    no_space_count = 0
    ticket_list = []
    today = datetime.datetime.today()
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    cmd = 'sh /infomity-mainte/bin/stcheck/KimSystemini_chk.sh | ' + DBST_EXCLUDE 
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    for x in range(len(teststr)):
        str1 = re.split(",", teststr[x])
        str2 = string.split(str1[2])
        kmidate = str_to_date3(str2[0])
        today = datetime.datetime.now()
        today = str_to_date3(str(today.year)+"/"+str(today.month)+"/"+str(today.day))
        diff_days = abs((today - kmidate).days)
        if diff_days > 4:
                ticket = de_classes.Ticket()
                ticket.customer_dbid = str1[0] + str1[1]
                ticket_list.append(ticket)
                i = i + 1
    return ticket_list
		
def get_new_tickest_from_db(server, exclude_list):
    i = 0
    no_space_count = 0
    ticket_list = []
    today = datetime.datetime.today()
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    cmd = 'cd /usr/local/pgsql/data/infomity-mainte/log/; cat `ls -tr infomity-userDBRepliDelayPingCSV-ForStatistics-' + str(today.year) + '* | tail -n 1` | '+exclude_list+'| grep ",[0-1]"'
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    for x in range(len(teststr)):
        str1 = re.split("\|", teststr[x])
        if len(str1) == 3:
            str2 = re.split(":", str1[2])
            if int(str2[0]) > 100:
                ticket = de_classes.Ticket()
                ticket.customer_dbid = str1[1]
                ticket_list.append(ticket)
                i = i + 1
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
            service.contract_end = str2[0]

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
    fmts = ['Text','m/D/YYYY',]

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
    print "tickets lenght = " + str(len(tickets))    
    for ticket in tickets:	
        if ticket.image_pilot_version[:3] == "1.7":
            if len(str(ticket.ticket_close_date)) > 1:
                cur_worksheet = worksheet_closed_dbdb9a
                cur_row = rownum_closed_dbdb9a
                rownum_closed_dbdb9a +=1
            else:
                cur_worksheet = worksheet_dbdb9a
                cur_row = rownum_dbdb9a
                rownum_dbdb9a +=1
        else: 
            if len(str(ticket.ticket_close_date)) > 1:
                cur_worksheet = worksheet_closed_dbdb
                cur_row = rownum_closed_dbdb
                rownum_closed_dbdb +=1
            else:
                cur_worksheet = worksheet_dbdb
                cur_row = rownum_dbdb
                rownum_dbdb +=1
		
        if len(ticket.author) <= 1:
			ticket.author = "Evghenii"

        cur_worksheet.write(cur_row, 1, "", style=style_general)
        #only closed tickets has close date
        print ticket.ticket_open_date
        """
        if type(ticket.ticket_close_date) is datetime.datetime and len(ticket.ticket_close_date.strftime("%B,%d,%Y")) > 1:
            cur_worksheet.write(cur_row, 2, ticket.ticket_close_date.strftime('%m/%d/%Y'), style=style_date)
        else:
            cur_worksheet.write(cur_row, 2, "", style=style_general)
        """
        cur_worksheet.write(cur_row, 3, ticket.ticket_open_date.strftime('%m/%d/%Y'), style=style_date)
        for x in range(4, 26):
            cur_worksheet.write(cur_row, x, "", style=style_general)
        """
		cur_worksheet.write(cur_row, 4, ticket.author, style=style_general)
        cur_worksheet.write(cur_row, 5, ticket.customer_name, style=style_general)
        cur_worksheet.write(cur_row, 6, ticket.customer_dbid, style=style_general)
        cur_worksheet.write(cur_row, 7, ticket.manual_synchronization, style=style_general)
        if isinstance(ticket.support_date_end, basestring):
            ticket.support_date_end = str_to_date4(ticket.support_date_end)
        cur_worksheet.write(cur_row, 8, ticket.support_date_end.strftime('%m/%d/%Y'), style=style_date)
        if isinstance(ticket.contract_date_start, basestring):
            ticket.contract_date_start = str_to_date4(ticket.contract_date_start)
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
        cur_worksheet.write(cur_row, 19, ticket.comment, style=style_general)
        cur_worksheet.write(cur_row, 20, ticket.action1_db, style=style_general)
        cur_worksheet.write(cur_row, 21, ticket.action1_request, style=style_general)
        cur_worksheet.write(cur_row, 22, ticket.action2_st, style=style_general)
        
        if not isinstance(ticket.action2_db, basestring):
            cur_worksheet.write(cur_row, 23, ticket.action2_db.strftime('%m/%d/%Y'), style=style_date)
        
        cur_worksheet.write(cur_row, 24, ticket.action3_request, style=style_general)
        cur_worksheet.write(cur_row, 26, ticket.call_center_comment, style=style_general)
        
        if len(ticket.services) > 0:
            for serv in ticket.services:		
                if serv.service_name == "DATABANK":
                    d1 = str_to_date4(serv.service_end[:7])
                    cur_worksheet.write(cur_row, 8, d1.strftime('%Y/%m/%d/'), style=style_date)
                    d1 = str_to_date4(serv.contract_start[:7])
                    cur_worksheet.write(cur_row, 9, d1.strftime('%m/%d/%Y'), style=style_date)
		"""	
	wb.save(EXCELFILE2)
	#os.remove(EXCELFILE)
	#os.rename(EXCELFILE2, EXCELFILE)

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
                        ticket.kim_system_ini_timestamp = str_to_date3(str1[2][:3])
                elif type(ticket.kim_system_ini_timestamp) is datetime.datetime:
                    if len(ticket.kim_system_ini_timestamp.strftime("%B,%d,%Y")) < 1:
                        ticket.kim_system_ini_timestamp = str_to_date3(str1[2][:3])

        ticket_list.append(ticket)
    return ticket_list
		
def get_db_repli(server, ticket):
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    cmd="/usr/local/pgsql/DBRepli/DBRepliDelay.sh " + ticket.customer_dbid
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    ticket.manual_synchronization = "NG"
    if len(teststr) == 1 and len(teststr[0]) > 1:
        ticket.tmp = teststr[0]
        if teststr[0] == "DB DOWN":
            ticket.manual_synchronization = "NG"
        else:
            str1 = re.split(",", teststr[0])
            str2 = re.split(":", teststr[0])
            if int(str1[1]) == 1 and int(str2[0]) > 100:
			    ticket.manual_synchronization = "NG"
            else:
                ticket.manual_synchronization = "OK"
    return ticket
		
def get_another_database(server, ticket):
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    cmd="/usr/local/pgsql/bin/psql -U postgres -t -p 5433 postgres -c \"SELECT datname FROM pg_database where datname like '"+ticket.customer_dbid[:7]+"%';\""
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    if len(teststr) >= 2:
        ticket.another_db_exist = teststr[1]
        ticket.another_db_status = "NG"
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
    cmd = "/usr/local/pgsql/bin/psql -U postgres -p 5433 " + ticket.customer_dbid[:7] + " -t -c \"select st_last_event_ts from _c1.sl_status ;\""
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    if len(teststr) > 0:
        ticket.last_db_access_date = str_to_date1(teststr[0])
	
    cmd = "/usr/local/pgsql/bin/psql -U postgres -p 5433 " + ticket.customer_dbid[:7] + " -t -c \"select st_last_received_event_ts from _c1.sl_status ;\""
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    if len(teststr) > 0:
        ticket.last_db_sync_date = str_to_date1(teststr[0])
	
    return ticket

def fill_new_tickets(ticket):
    #fill Customer name
    CUSTOMER_NAME="nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
    str = exec_ldap_query(LDAPSRV, CUSTOMER_NAME, "(objectClass=*)", "medicalInstitutionName")
    if len(str) > 0:
        ticket.customer_name = str[0]
		
	#Fill cussomer phone number	
    TEL_NUMBER="nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
    str = exec_ldap_query(LDAPSRV, TEL_NUMBER,"(objectClass=*)", "telephoneNumber")
    if len(str) > 0:
        ticket.phone_number = str[0]
    
    PACKAGE_INFO = "o=packageContracts,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
    str = exec_ldap_query(LDAPSRV, PACKAGE_INFO,"contractFlag=TRUE", "infomityServicePackageCode")
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
	
    for serv in ticket.services:
        if serv.service_name == "DATABANK":
            ticket.support_date_end = serv.service_end
            ticket.contract_date_start = serv.contract_start

    return ticket
	
def fill_and_write_new_tickets(workbook, tickets):
    tickets = get_client_version(DBSTSRV, tickets)
    
    for ticket in tickets:
        ifCreateionDate = False
        if isinstance(ticket.ticket_open_date, basestring):
            if len(ticket.ticket_open_date) < 1:
                ifCreateionDate = True         
        elif type(ticket.ticket_open_date) is datetime.datetime:
            if len(ticket.ticket_open_date.strftime("%B,%d,%Y")) < 1:
                ifCreateionDate = True
        
        if ifCreateionDate is True:
            ticket = fill_new_tickets(ticket)
            print ticket.customer_dbid + "   " + ticket.image_pilot_version[:8]

    #tickets = tickets.sort(key=lambda item:['ticket_open_date'], reverse=True)
    #tickets = sorted(tickets, key=itemgetter(2), reverse=True)

    rewrite_workbook(workbook, tickets)

def check_if_replicated(server, tickets):
    ticket_list = []    
    today = datetime.datetime.today()
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    cmd = "sh /infomity-mainte/bin/stcheck/KimSystemini_chk.sh "
    stdin, stdout, stderr = ssh.exec_command(cmd)
    teststr = stdout.readlines()
    for x in range(len(teststr)):
        str1 = re.split(",", teststr[x])
        str2 = string.split(str1[2])
        kmidate = str_to_date3(str2[0])
        for ticket in tickets:
            if str1[0] == ticket.customer_dbid[:7] and str1[1] == ticket.customer_dbid[7:]:   
                today = datetime.datetime.now()
                today = str_to_date3(str(today.year)+"/"+str(today.month)+"/"+str(today.day))
                diff_days = abs((today - kmidate).days)
                if diff_days < 4:
                    ticket.customer_dbid = str1[0] + str1[1]
					#cheking if DB well replicated
                    if ticket.image_pilot_version[:3] == "1.7":
                        ticket = get_db_repli(DBDB9ASRV, ticket)
                    else:
                        ticket = get_db_repli(DBDBSRV, ticket)
                    if ticket.manual_synchronization == "OK":
                        ticket_close_date = datetime.datetime.today()
                        ticket_list.append(ticket)

    return ticket_list


def print_date(function_name):
    d = datetime.datetime.today()
    print str(d) + " " + function_name

#read excel file
print_date('begin')

xl_workbook = xlrd.open_workbook(EXCELFILE, formatting_info=True)	

#closed_dbdblist = de_read_sheet(xl_workbook, 'Closed-dbdb', 3)
#print_date('de_utils.de_utils.de_read_sheet')

#closed_dbdb9alist = de_read_sheet(xl_workbook,'Closed-dbdb9a', 3)
#print_date('de_utils.de_utils.de_read_sheet')

dbdblist = de_read_sheet(xl_workbook,'dbdb', 11)
print_date('de_utils.de_utils.de_read_sheet')
print "dbdblist lenght = " + str(len(dbdblist))
#dbdb9alist = de_read_sheet(xl_workbook,'dbdb9a', 11)
#print_date('de_utils.de_utils.de_read_sheet')

#close replicated tickets
#dbdblist_replicated = check_if_replicated(DBSTSRV, dbdblist)
#print_date('check_if_replicated')

#dbdb9alist_replicated = check_if_replicated(DBSTSRV, dbdb9alist)
#print_date('check_if_replicated')

#dbdblist = remove_dublicates(dbdblist, dbdblist_replicated)
#print_date('remove_dublicates')

#dbdb9alist = remove_dublicates(dbdb9alist, dbdb9alist_replicated)
#print_date('remove_dublicates')

#closed_dbdblist = assembly_two_list(closed_dbdblist, dbdblist_replicated)
#print_date('assembly_two_list')

#closed_dbdb9alist = assembly_two_list(closed_dbdb9alist, dbdb9alist_replicated)
#print_date('assembly_two_list')

#check new tickets on servers
dbdb_new = get_new_tickest_from_db(DBDBSRV, DBDB_EXCLUDE)
print_date('get_new_tickest_from_db')
print "dbdb_new lenght = " + str(len(dbdb_new))
#dbdb9a_new = get_new_tickest_from_db(DBDB9ASRV, DBDB9A_EXCLUDE)
#print_date('get_new_tickest_from_db')

dbst_new = get_new_tickest_from_st(DBSTSRV)
print_date('get_new_tickest_from_st')

dbdb_new = remove_dublicates(dbdb_new, dbdblist)
print_date('remove_dublicates')
print "dbdb_new lenght = " + str(len(dbdb_new))
#dbdb9a_new = remove_dublicates(dbdb9a_new, dbdb9alist)
#print_date('remove_dublicates')

#dbst_new = remove_dublicates(dbst_new, dbdb9alist)
#print_date('remove_dublicates')
#dbst_new = remove_dublicates(dbst_new, dbdb9a_new)
#print_date('remove_dublicates')

common_list = assembly_two_list(dbdblist, dbdb_new)
print "common_list lenght = " + str(len(common_list))
#common_list = assembly_two_list(dbst_new, dbdb9a_new)
#print_date('assembly_two_list')
#common_list = assembly_two_list(common_list, dbst_new)
#print_date('assembly_two_list')
#common_list = assembly_two_list(common_list, closed_dbdblist)
#print_date('assembly_two_list')
#common_list = assembly_two_list(common_list, closed_dbdb9alist)
#print_date('assembly_two_list')


fill_and_write_new_tickets(xl_workbook, common_list)
print_date('fill_and_write_new_tickets')

