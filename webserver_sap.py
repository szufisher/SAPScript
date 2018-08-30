# tips
# vba vs python: Key word capital letter, If -> if, Else If-> elif, Exit For ->break, On Error Resume Next -> try except pass
# sendVKey 0 ->sendVKey(0), .press ->.press()
# dependent module pywin32, pymssql
# configparser with default, otherwise there is error when config key not available
# exe by pyinstaller, 1.if there is error, when double click to run the exe, the cmd window disappeared instantly, no chance to see error log
# run cmd, cd to the exe file folder, type the file name, then cmd window will show the error
# ImportError: No module named decimal, it is due to mssql dynamicly import this module, pyinstaller does not packed it into exe, 
# import in your py file even it is not used directly in the py file!

from cgi import parse_qs
import ConfigParser,decimal
from wsgiref.simple_server import make_server
import sys, win32com.client
import pymssql

myDefaults = {'db_server':'','db_name':'','web_port':'8080'}
config = ConfigParser.ConfigParser(defaults=myDefaults)
config.read('webserver_sap.ini')

# -Sub Main--------------------------------------------------------------
def start_sap():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)        
        session = connection.Children(0)
        return session
    except:
        print(sys.exc_info()[0])

def close_sap(session):
    session = None
    connection = None
    application = None
    SapGuiAuto = None

def execute_transaction(session,po, po_item, po_item_qty):   
    col_item_ok = 0
    CurrentRow = 0 
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmigo"
    session.findById("wnd[0]").sendVKey(0)
    print 'session started'    
    try:
        session.findById("wnd[0]/usr").FindByNameEx("BUTTON_HEADER_TOGGLE", 40).press()
    except:
        pass    
    try:
        session.findById("wnd[0]/usr").FindByNameEx("BUTTON_ITEMDETAIL", 40).press()
    except:
        pass
    session.findById("wnd[0]/usr").FindByNameEx("GODYNPRO-ACTION", 34).Key = "A01"
    session.findById("wnd[0]/usr/").FindByNameEx("GODYNPRO-REFDOC", 34).Key = "R01"
    session.findById("wnd[0]/usr/").FindByNameEx("GODYNPRO-PO_NUMBER", 32).Text = po
    session.findById("wnd[0]/usr/").FindByNameEx("GODYNPRO-PO_ITEM", 31).Text = po_item
    session.findById("wnd[0]/usr/").FindByNameEx("GODYNPRO-PO_WERKS", 32).Text = ""
    session.findById("wnd[0]/usr").FindByNameEx("GODEFAULT_TV-BWART", 32).Text = "101"
    session.findById("wnd[0]").sendVKey(0)
    mess_type = session.findById("wnd[0]/sbar").MessageType
    if mess_type == "W":
        session.findById("wnd[0]").sendVKey(0)
    elif mess_type == "E":
        error_message = session.findById("wnd[0]/sbar").Text
    result = session.findById("wnd[0]/sbar").Text
    if result != "":
        return "Failed", result
    # columns in different row has different column index due to the fact that some columns not present in some rows such as the 5th documentation column
    if col_item_ok == 0:
        cur_row = session.findById("wnd[0]/usr").FindByNameEx("SAPLMIGOTV_GOITEM", 80).Rows(CurrentRow)
        col_count = 0
        for ii  in range(0, cur_row.Count - 1):
			if cur_row(ii + 0).Name == "GOITEM-TAKE_IT":
				col_item_ok = ii + 0
				col_count = col_count + 1
			elif cur_row(ii + 0).Name == "GOITEM-ERFMG":
				col_mat_qty = ii + 0
				col_count = col_count + 1
			if col_count > 1:
				break
        cur_row = None
    session.findById("wnd[0]/usr").FindByNameEx("SAPLMIGOTV_GOITEM", 80).Rows(CurrentRow)( col_mat_qty + 0).Text = po_item_qty
    session.findById("wnd[0]/usr").FindByNameEx("SAPLMIGOTV_GOITEM", 80).Rows(CurrentRow)(col_item_ok + 0).Selected = True

    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    MessageType = session.findById("wnd[0]/sbar").MessageType
    if MessageType == "W":
        session.findById("wnd[0]").sendVKey(0)
    result = session.findById("wnd[0]/sbar").Text
    if "Material document " in result:
        result = result.split()[2]
        return "OK", result
    else:
        return "Failed", result

def connect_db():
	db_server = config.get('DEFAULT','db_server')
	db_name = config.get('DEFAULT','db_name')
	conn= pymssql.connect(server=db_server,database=db_name)
	cursor = conn.cursor()
	return conn, cursor

def get_from_db(cursor, gr):	
	sql = """select c_PONo,c_ItemNo,c_ReceivingQuantity from app_fd_P02_SelectedGREvent
		where (c_InputStatus is Null or c_InputStatus<>'OK') and c_GRProcessIDNo = '%s'""" % gr
	#print sql
	cursor.execute(sql)
	result = [row for row in cursor]
	return result
	
def update_to_db(cursor, gr, po, po_item, status, message):	
	cursor.execute("""update app_fd_P02_SelectedGREvent set c_InputStatus = '%s' ,c_Message= '%s'
		where c_GRProcessIDNo='%s' and c_PONo = '%s' and c_ItemNo='%s'""" % (status, message, gr, po, po_item))
	return True
	
def close_db(conn):
	conn.close()
	
def simple_app(environ, start_response):
    status = '200 OK'
    headers = [('Content-Type', 'text/plain')]
    start_response(status, headers)
    if environ['REQUEST_METHOD'] == 'POST':
        request_body_size = int(environ.get('CONTENT_LENGTH', 0))
        request_body = environ['wsgi.input'].read(request_body_size)
        d = parse_qs(request_body)  # turns the qs to a dict
        return 'From POST: %s' % ''.join('%s: %s' % (k, v) for k, v in d.iteritems())
    else:  # GET
        d = parse_qs(environ['QUERY_STRING'])  # turns the qs to a dict
        gr = d.get('gr')
        po = d.get('po')
        item = d.get('item')
        item_qty = d.get('qty') or 1
        if gr:
			session = start_sap()
			messages=[]
			if session:
				conn, cursor = connect_db()				
				po_items = get_from_db(cursor, gr[0])
				if po_items:	
					for po_item in po_items:
						po,item,item_qty = po_item
						status, message = execute_transaction(session, po, item,item_qty)
						update_to_db(cursor, gr, po, item, status, message)
						messages.append('%s,%s' %(status,message))
				else:
					messages.append("Failed, gr not exist or already processed!")
				close_sap(session)
				close_db(conn)
			else:
				messages.append("Failed, No sap session")
			result = '\n'.join(messages)			
			print result
			return result.encode('ascii') if isinstance(result, unicode) else result
        elif po and item:
			print 'po: %s; item: %s' %(po, item)
			session = start_sap()
			if session:
				status, message = execute_transaction(session, po[0], item[0],item_qty[0])
				close_sap(session)
			else:
				status, message = "Failed", "No sap session"
			result = '%s,%s' %(status, message)						
			print result
			return result.encode('ascii') if isinstance(result, unicode) else result
        else:
            return 'URL should look like this: http://localhost:8080/?po=4700307581;item=10;qty=1'

def main():
	port = config.get('DEFAULT','web_port')
	httpd = make_server('', int(port), simple_app)
	print "Serving on port %s..." % port
	httpd.serve_forever()
	
if __name__ == "__main__":
  main()
