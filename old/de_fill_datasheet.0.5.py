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
import string
import ldaphelper
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

DBSTSRV = '10.227.9.205'
DBDBSRV = '10.227.9.219'
DBDB9ASRV = '10.227.9.223'
LDAPSRV = "ldap://10.227.9.233"
DBDBSUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb.sh'
DBDB9ASUMMARY = '/Admin/scripts/Summary-combined-cnt-report-dbdb9a.sh'
EXCELFILE = '04-03_Replication_failure_control_doc.xls'
DBDB_EXCLUDE = 'egrep -iv \"2000010|2008977|2004785|2008195|2009247|2009275|2009612|2008203|2008261|2008470|delay|DEMO|DBRepli\"'
DBDB9A_EXCLUDE = 'egrep -iv \"2000010|2008977|2004785|2008195|2009247|2009275|2009612|2008203|2008261|2012410|2012412|2012413|2012414|2012415|2012517|delay|DEMO|DBRepli\"'
DBST_EXCLUDE = 'egrep -iv \"1111111|2000010|2008977|2004785|2008195|2009247|2009275|2009612|2008203|2008261|2008470|DEMO|2012410|2012412|2012413|2012414|2012415|2012517|_2009451_backup20130926|H1USTEST|KMMG|mhustest|KimSystem.ini|backup_|2012102|2012409|2012411|2005040,59023265040A10102,|2005345,5902300617016024,|2007524,5902320573173975,|2007659,590232001139,|2007885,59023186645916137,|2008565,590232300001,|2008609,5902336241924962,|2008770,5902326130165507,|2008779,5902326519102819,|2008816,5902330176455434,|2008834,5902326308408561,|2008880,5902330594293755,|2008887,5902330800635855,|2008927,5902326801068366,|2008955,5902326235214421,|2008968,5902326906255933,|2008972,5902330168393170,|2008973,5902330628987942,|2009050,5902326247380041,|2009101,5902326281050326,|2009102,5902326515329586,|2009124,5902326323067626,|2009125,5902326644898398,|2009126,5902326965607020,|2009133,590236981485420,|2009162,5902326390764034,|2009197,5902326550614006,|2009439,5902330996320817,|2009451,5902330833817731,|2009451,5902334110896529,|2009676,59023300001,|2009676,5902330479449308,|2009749,5902330691271287,|2009790,5902330504560097,|2009790,5902330930792648,|2010077,590233456509572,|2010222,5902334153244039,|2010516,5902319108040445,|2010529,5902336713833534,|2010830,590233400001,|2009288,5902330351771896,|2009366,5902330808628696,|2008113,59020131680,|2008107,5902320PL12UA04819WW,|2001680,590230099C0676,|2009070,5902326489260566,|2007960,590231901455,|2002755,5902300MJ00717,|2009543,5902323956247757,|2009642,5902330105810346,|2008031,5902319374978909,|2008971,5902330358998116,|2009518,5902330860249699,|2008667,5902319783501179,|2008946,5902330707669365,|2008840,5902326494911924,|2009576,59023262UA1290GV0,|2008189,590231901732\"'

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
    package_info = ''
    phone_number = ''
    services = []
    products = []
	
class Service:
    service_code = ''
    service_name = ''
    contract_start = ''
    contract_end = ''
    service_start = ''
    service_end = ''
    contract_flag = ''
    full_size = ''
    empty_size = ''
    contract_pc_num = ''
    set_pc_num = ''

class Product:
    product_number = ''
    serial_number = ''
    empty_size = ''
    full_size = ''
    dbrepli_ip = ''
    storage_ip = ''
    stop_flag = ''
    create_time = ''
	
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
	
#converts string 2015/12/31 to date
def str_to_date3(value):
    format = "%Y/%m/%d"
    try:
	    d = datetime.datetime.strptime(value, format)
    except ValueError:
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

def parse_ldap_single_value(str):
	if len(str) > 0:
		try:
			value = int(str)
			print value
			return value
		except	ValueError, e:
			print str
			return str			

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
                ticket = Ticket()
                ticket.customer_dbid = str1[0] + str1[1]
                ticket.image_pilot_version = str1[3]
                ticket_list.append(ticket)
                i = i + 1
    return ticket_list
		
		
def get_new_tickest_from_db(server):
    i = 0
    no_space_count = 0
    ticket_list = []
    today = datetime.datetime.today()
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(server, username='root', password='konicaminolta')
    if server == '10.227.9.219':
        cmd = 'cd /usr/local/pgsql/data/infomity-mainte/log/; cat `ls -tr infomity-userDBRepliDelayPingCSV-ForStatistics-' + str(today.year) + '* | tail -n 1` | '+DBDB_EXCLUDE+'| grep ",[0-1]"'
    else: 	
        cmd = 'cd /usr/local/pgsql/data/infomity-mainte/log/; cat `ls -tr infomity-userDBRepliDelayPingCSV-ForStatistics-' + str(today.year) + '* | tail -n 1` | '+DBDB9A_EXCLUDE+'| grep ",[0-1]"'
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

def assemply_trouble_list(list1, list2, list3):
    list_tmp = []
    for x in range(len(list1)):
        list_tmp.append(list1[x])
    for x in range(len(list2)):
        list_tmp.append(list2[x])
    for x in range(len(list3)):
        list_tmp.append(list3[x])
    return list_tmp

def try_ldap(searchstr, searchFilter, base):
    values = []
    BIND_DN = "cn=Manager,dc=infomity,dc=net"
    BIND_PASS = "secret"
    attributeFilter = [base]
    try:
        ldap.set_option(ldap.OPT_X_TLS_REQUIRE_CERT, 0)
        lcon_emg = ldap.initialize(LDAPSRV)
        lcon_emg.simple_bind_s(BIND_DN, BIND_PASS)
        ldap_result_id = lcon_emg.search_s(searchstr, ldap.SCOPE_SUBTREE, searchFilter, attributeFilter)
        
        res = ldaphelper.get_search_results(ldap_result_id)
        for i in range(len(res)):
            str = res[i].pretty_print2()
            if len(str) > 0:
                values.append(str)
    except ldap.LDAPError, e:
        pass
    return values	
	
def try_ldap2(basedn , searchFilter, base):
    values = []
    BIND_DN = "cn=Manager,dc=infomity,dc=net"
    BIND_PASS = "secret"
    attributeFilter = [base]
    try:
        ldap.set_option(ldap.OPT_X_TLS_REQUIRE_CERT, 0)
        lcon_emg = ldap.initialize(LDAPSRV)
        lcon_emg.simple_bind_s(BIND_DN, BIND_PASS)
        ldap_result_id = lcon_emg.search_s(basedn, ldap.SCOPE_SUBTREE, searchFilter, attributeFilter)
        
        res = ldaphelper.get_search_results(ldap_result_id)
        for i in res:
            str = i.pretty_print2()
            if str is not None and len(str) > 0:
                values.append(str)
    except ldap.LDAPError, e:
        pass	
    return values	

def fill_services_list(siteid):
    services_list = []
    SERVICE_CONTRACTS = "o=contracts,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
    services = try_ldap(SERVICE_CONTRACTS, "(contractFlag=TRUE)", "infomityServiceCode")

    for serv in services:
        print serv
        service = Service()
        str1 = "infomityServiceCode=" + serv + ",o=services,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
        str2 = try_ldap2(str1, "(objectClass=*)", "infomityServiceName")
        if len(str2) > 0:
            service.service_name = str2[0]
            #print str2[0]
        
        str1 = "infomityServiceCode=" + serv + ",o=contracts,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
        str2 = try_ldap2(str1, "(objectClass=*)", "contractConclusionDay")
        if len(str2) > 0:
            service.contract_start = str2[0]
            #print str2[0]

        str1 = "infomityServiceCode=" + serv + ",o=contracts,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
        str2 = try_ldap2(str1, "(objectClass=*)", "contractFinishDay")
        if len(str2) > 0:
            service.contract_end = str2[0]
            #print str2[0]

        str1 = "infomityServiceCode=" + serv + ",o=contracts,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
        str2 = try_ldap2(str1, "(objectClass=*)", "infomityServiceStart")
        if len(str2) > 0:
            service.service_start = str2[0]
            #print str2[0]

        str1 = "infomityServiceCode=" + serv + ",o=contracts,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
        str2 = try_ldap2(str1, "(objectClass=*)", "infomityServiceEnd")
        if len(str2) > 0:
            service.service_end = str2[0]
            #print str2[0]
	
        str1 = "infomityServiceCode=" + serv + ",o=contracts,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
        str2 = try_ldap2(str1, "(objectClass=*)", "contractFlag")
        if len(str2) > 0:
            service.contract_flag = str2[0]
            #print str2[0]
		
        str1 = "cn=InstCapacity,o=settings,infomityServiceCode=MGBOX,o=services,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
        str2 = try_ldap2(str1, "(objectClass=*)", "maximum")
        if len(str2) > 0:
            service.full_size = str2[0]
            #print str2[0]

        str1 = "cn=mboxEmptySize,o=settings,infomityServiceCode=MGBOX,o=services,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
        str2 = try_ldap2(str1, "(objectClass=*)", "numericValue")
        if len(str2) > 0:
            service.empty_size = str2[0]
            #print str2[0]		

        str1 = "cn=databankContractPCNumber,o=settings,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
        str2 = try_ldap2(str1, "(objectClass=*)", "numericValue")
        if len(str2) > 0:
            service.contract_pc_num = str2[0]
            #print str2[0]
	
        str1 = "cn=databankSetPCNumber,o=settings,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
        str2 = try_ldap2(str1, "(objectClass=*)", "numericValue")
        if len(str2) > 0:
            service.set_pc_num = str2[0]
            #print str2[0]
        
        services_list.append(service)
    
    return services_list
	

def fill_products_list(siteid):
    products_list = []
    PRODUCTS = "o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
    products = try_ldap(PRODUCTS, "(objectClass=*)", "productNumber")

    for prod in products:
        SERIAL = "o=devices,productNumber=" + prod + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
        serials = try_ldap(PRODUCTS, "(objectClass=*)", "serialNumber")
        for serl in serials:
            product = Product()
            product.product_number = prod
            product.serial_number = serl
            #print prod
            #print serl
            str1 = "cn=backupEmptySize,serialNumber=" + serl + ",o=devices,productNumber=" + prod + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
            str2 = try_ldap2(str1, "(objectClass=*)", "numericValue")
            if len(str2) > 0:
                product.empty_size = str2[0]
                #print str2[0]

            str1 = "cn=backupFullSize,serialNumber=" + serl + ",o=devices,productNumber=" + prod + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
            str2 = try_ldap2(str1, "(objectClass=*)", "numericValue")
            if len(str2) > 0:
                product.full_size = str2[0]
                #print str2[0]

            str1 = "cn=useDBRepliIPAddress,serialNumber=" + serl + ",o=devices,productNumber=" + prod + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
            str2 = try_ldap2(str1, "(objectClass=*)", "value")
            if len(str2) > 0:
                product.dbrepli_ip = str2[0]
                #print str2[0]

            str1 = "cn=useStorageIPAddress,serialNumber=" + serl + ",o=devices,productNumber=" + prod + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
            str2 = try_ldap2(str1, "(objectClass=*)", "value")
            if len(str2) > 0:
                product.storage_ip = str2[0]
                #print str2[0]

            str1 = "cn=backupStopFlag,serialNumber=" + serl + ",o=devices,productNumber=" + prod + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
            str2 = try_ldap2(str1, "(objectClass=*)", "enableFlag")
            if len(str2) > 0:
                product.stop_flag = str2[0]
                #print str2[0]
				
            str1 = "serialNumber=" + serl + ",o=devices,productNumber=" + prod + ",o=products,infomityServiceCode=DATABANK,o=services,nettsInstitutionCode=" + siteid + ",o=institutions,dc=infomity,dc=net"
            str2 = try_ldap2(str1, "(objectClass=*)", "createTimestamp")
            if len(str2) > 0:
                product.create_time = str2[0]
                #print str2[0]
				
            products_list.append(product)
    
	return products_list
	
def fill_new_tickets(tickests):
    for ticket in tickests:
        USER_BASE = "nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        TEL_NUMBER="nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        str = try_ldap2(TEL_NUMBER,"contractFlag=TRUE", "telephoneNumber")
        if len(str) > 0:
            ticket.phone_number = str[0]
            print str[0]

        PACKAGE_INFO = "o=packageContracts,nettsInstitutionCode=" + ticket.customer_dbid[:7] + ",o=institutions,dc=infomity,dc=net"
        str = try_ldap2(PACKAGE_INFO,"contractFlag=TRUE", "infomityServicePackageCode")
        if len(str) > 0:
            ticket.package_info = str[0]
            print str[0]

        ticket.services = fill_services_list(ticket.customer_dbid[:7])
        ticket.pruducts = fill_products_list(ticket.customer_dbid[:7])
	
xl_workbook = xlrd.open_workbook(EXCELFILE)	

closed_dbdblist = de_read_sheet('Closed-dbdb', 3)
print "de_read_sheet finished"

closed_dbdb9alist = de_read_sheet('Closed-dbdb9a', 3)
print "de_read_sheet finished"

dbdblist = de_read_sheet('dbdb', 11)
print "de_read_sheet finished"

dbdb9alist = de_read_sheet('dbdb9a', 11)
print "de_read_sheet finished"

dbdb_new = get_new_tickest_from_db(DBDBSRV)
print "get_new_tickest_from_db finished"

dbdb9a_new = get_new_tickest_from_db(DBDB9ASRV)
print "get_new_tickest_from_db finished"

dbst_new = get_new_tickest_from_st(DBSTSRV)
print "get_new_tickest_from_st finished"

dbdb_new = remove_dublicates(dbdb_new, dbdblist)
print "remove_dublicates finished"

dbdb9a_new = remove_dublicates(dbdb9a_new, dbdb9alist)
print "remove_dublicates finished"

dbst_new = remove_dublicates(dbst_new, dbdblist)
print "remove_dublicates finished"

dbst_new = remove_dublicates(dbst_new, dbdb9alist)
print "remove_dublicates finished"

dbst_new = remove_dublicates(dbst_new, dbdb9a_new)
print "remove_dublicates finished"

common_list = assemply_trouble_list(dbst_new, dbdb9a_new, dbst_new)
print "assemply_trouble_list finished"

fill_new_tickets(common_list)
print "fill_new_tickets finished"


		
