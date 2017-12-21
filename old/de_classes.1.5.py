class Ticket:
    ticket_number = ''
    ticket_close_date = ''
    ticket_open_date = ''
    author = ''
    customer_name = ''
    customer_dbid = ''
    manual_synchronization = '' #May be OK or NG
    support_date_end = ''
    contract_date_start = ''
    another_db_exist = ""
    another_db_status = ""
    table_exist_check = '' #OK, NULL or DOWN
    last_db_access_date = ''
    last_db_sync_date = ''
    last_db_sync_hours = '' #how much hours passed from last sync
    log_check = ''
    st_server = 'ST1' #always ST1
    kim_system_ini_timestamp = ''
    image_pilot_version = ''
    comment = ''
    action1_db = ''#column 20
    action1_request = ''#column 21
    comment_out_from_st = ''#column 22 action2_st
    comment_out_from_pg_hba = ''#column 23 action2_db
    action3_request = ''#column 24
    call_center_comment = '' #column 25
    package_info = ''
    phone_number = ''
    sheet_name = ''
    databank_enabled = True
    contract_expired = ''
    services = []
    products = [] #collecting only for DATABANK service
	
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
