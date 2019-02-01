#VBA=>PY
# sendvkey 0 => senvkey(0)
# .press =>.press()
# .select => .select()
# on error resume next => try   except
# set obj = session.xxx => obj = session.xxx


# -*- coding: utf-8 -*-
#from cgi import papo_linee_qs
#import ConfigPapo_lineer, decimal
from wsgiref.simple_server import make_server
import sys, win32com.client
import pymssql

myDefaults = {'db_server':'xxx','db_name':'xxx','web_port':'8080'}
#config = ConfigPapo_lineer.ConfigPapo_lineer(defaults=myDefaults)
#config.read('webserver_sap.ini')

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

def execute_transaction(session):
    po_lines = [['10', '9300435', 'AC8', '10452233', u'test python 中文', 'aba', 'k', '25.5', 'pc', '15.42',
                 'USD','30r1', '050319', 'yuxin', '61790050', 'wh', 'lipingping', '', '','', 
                 '830702', '100', '', '','1', '', '001'],
                ['20', '9300435', 'AC8', '10452234', 'test python', 'aba', 'k', '28.5', 'pc', '25.42',
                  'USD','30r1','050319', 'yuxin', '61790050', 'wh', 'lipingping', '', '', '',
                  '830701', '100', '','', '1', '', '001'],
				  ['30', '9300435', 'AC8', '10452233', u'test更多中文 python', 'aba', 'k', '25.5', 'pc', '15.42',
                 'USD','30r1', '050319', 'yuxin', '61790050', 'wh', 'lipingping', '', '','', 
                 '830702', '60', '', '','1', '', '001'],
                ['30', '9300435', 'AC8', '10452234', 'test更多中文 python', 'aba', 'k', '28.5', 'pc', '25.42',
                  'USD','30r1','050319', 'yuxin', '61790050', 'wh', 'lipingping', '', '', '',
                  '830701', '40', '','', '1', '', '001']
                 ]
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme2n"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[16]").press()
    dyn_select = session.findById("wnd[0]/usr").FindByNameEx("shellcont[1]", 51).FindByNameEx("shell", 122)
    dyn_select.expandNode("          1")
    dyn_select.selectNode("         56")
    dyn_select.doubleClickNode("         56")
    #dyn_select = Nothing
    session.findById("wnd[0]/usr").FindByNameEx("%%DYN001-LOW", 31).Text = po_lines[0][-1]
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    msg = session.findById("wnd[0]/sbar").Text
    if not msg:
        old_po = ""
        try:
            session.findById("wnd[0]/usr/lbl[1,5]")
            old_po = session.findById("wnd[0]/usr/lbl[1,5]").Text
        except:
            print("PO " & old_po & " already created for this PA, please check PO header->communication->Your reference field")
            return
    # Set po_line = conn.Execute("select * from View_P03_PO where c_ReferID = " & "'" & po & "'")
    item_count = 0
    last_po_item = ""

    if not po_lines:
        return

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme21n"
    session.findById("wnd[0]").sendVKey(0)
    try:
        session.findById("wnd[0]/usr").FindByNameEx("DYN_4000-BUTTON", 40).press()
    except:
        pass
    session.findById("wnd[0]/usr").FindByNameEx("MEPO_TOPLINE-BSART", 34).Key = "NB"
    session.findById("wnd[0]/usr").FindByNameEx("MEPO_TOPLINE-SUPERFIELD", 32).Text = po_lines[0][1]
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(0)
    try:
		session.findById("wnd[0]/usr").FindByNameEx("MEPO1222-EKORG", 32).Text = '3000'
		session.findById("wnd[0]/usr").FindByNameEx("MEPO1222-BUKRS", 32).Text = 'CN10'
		session.findById("wnd[0]/usr").FindByNameEx("MEPO1222-EKGRP", 32).Text = po_lines[0][2]
    except:
        pass
    curr = po_lines[0][10]
    session.findById("wnd[0]/usr").FindByNameEx("TABHDT1", 91).Select()
    session.findById("wnd[0]/usr").FindByNameEx("MEPO1226-WAERS", 32).Text = curr
    session.findById("wnd[0]/usr").FindByNameEx("TABHDT5", 91).Select()
    session.findById("wnd[0]/usr").FindByNameEx("MEPOCOMM-UNSEZ", 31).Text = po_lines[0][-1]
    session.findById("wnd[0]").sendVKey(0)
    if session.findById("wnd[0]/sbar").messagetype == "E":
        print("Error with message: " & session.findById("wnd[0]/sbar").Text)
        return
    Ct = 0
    Position = 0
    for idx, po_line in enumerate(po_lines):
        po_item = po_line[0]
        vendor = po_line[1]
        # po_date = po_line[2]
        purchase_group = po_line[2].upper()
        material = po_line[3]
        mat_desc = po_line[4]
        if material:
            mat_desc = material + " " + mat_desc
        mat_desc = mat_desc[:39]
        # mat_desc = Replace(Replace(mat_desc, Chr(13), ""), Chr(10), "")  'remove carriage return/line break in the string!!
        material_group = po_line[5]
        aac = po_line[6].upper()
        quantity = po_line[7]
        uom = po_line[8]
        Price = po_line[9]
        plant = po_line[11]
        del_date = po_line[12]
        requisitioner = po_line[13][:11]
        gl_account = po_line[14]
        unloading_point = po_line[15]
        recipient = po_line[16]
        internal_order = po_line[17]
        wbs = po_line[18]
        asset = po_line[19]
        cost_center = po_line[20]
        percentage = po_line[21]
        tax_code = po_line[23]
        price_base = po_line[24]
        prno = po_line[26]
        if not material_group:
            material_group = "qsa"
        if po_item != last_po_item:
            last_po_item = po_item
            item_count += 1
            cost_split_count = 1
            row = 0
        else:
            cost_split_count += 1
            row = 1
        if cost_split_count == 1:
            cur_row = session.findById("wnd[0]/usr").FindByNameEx("SAPLMEGUITC_1211", 80).Rows(Ct)
            cur_row(2).Text = aac
            cur_row(5).Text = mat_desc
            cur_row(6).Text = quantity
            cur_row(7).Text = uom
            cur_row(9).Text = del_date
            cur_row(10).Text = Price
            cur_row(12).Text = price_base
            cur_row(14).Text = material_group
            cur_row(15).Text = plant
            cur_row(18).Text = prno
            cur_row(19).Text = requisitioner
            if Ct == 0:
                Ct = 1
            session.findById("wnd[0]").sendVKey(0)
        if aac == "K" and  percentage != "100":
            session.findById("wnd[0]/usr").FindByNameEx("TABIDT12", 91).Select()
            if row == 0:
                try:
                    session.findById("wnd[0]/usr").FindByNameEx("MEACCT1200-VRTKZ", 34).Key = "2"
                except:		
                    session.findById("wnd[0]/usr").FindByNameEx("MEACCT1200TB", 50).FindByNameEx("shell", 122).pressButton("MEAC1200DETAILTOGGLE")
                    session.findById("wnd[0]/usr").FindByNameEx("MEACCT1200-VRTKZ", 34).Key = "2"
                session.findById("wnd[0]/usr").FindByNameEx("MEACCT1200-TWRKZ", 34).Key = "2"
                session.findById("wnd[0]").sendVKey(0)
                first_percentage = percentage
            cur_row = session.findById("wnd[0]/usr").FindByNameEx("SAPLMEACCTVIDYN_1000TC", 80).Rows(row)
            cur_row(3).Text = percentage
            cur_row(4).Text = cost_center
            cur_row(5).Text = gl_account
            cur_row(7).Text = unloading_point
            cur_row(8).Text = recipient
            Position = session.findById("wnd[0]/usr").FindByNameEx("SAPLMEACCTVIDYN_1000TC", 80).verticalScrollbar.Position
            if cost_split_count == 2:
                session.findById("wnd[0]/usr").FindByNameEx("SAPLMEACCTVIDYN_1000TC", 80).verticalScrollbar.Position = 0
                cur_row = session.findById("wnd[0]/usr").FindByNameEx("SAPLMEACCTVIDYN_1000TC", 80).Rows(0)
                cur_row(3).Text = first_percentage
            session.findById("wnd[0]/usr").FindByNameEx("SAPLMEACCTVIDYN_1000TC", 80).verticalScrollbar.Position = Position + 1
        else:
            if aac in ["K","F","P","A"]:
                try:
                    unloading_field = session.findById("wnd[0]/usr").FindByNameEx("MEACCT1100-ABLAD", 31)
                except:
                    if not unloading_field:
                        session.findById("wnd[0]/usr").FindByNameEx("MEACCT1200TB", 50).FindByNameEx("shell", 122).pressButton("MEAC1200DETAILTOGGLE")
            session.findById("wnd[0]/usr").FindByNameEx("TABIDT12", 91).Select()
            session.findById("wnd[0]/usr").FindByNameEx("MEACCT1100-ABLAD", 31).Text = unloading_point
            session.findById("wnd[0]/usr").FindByNameEx("MEACCT1100-WEMPF", 31).Text = recipient
            if aac != "A":
                session.findById("wnd[0]/usr").FindByNameEx("MEACCT1100-SAKTO", 32).Text = gl_account
            if aac == "K":
                session.findById("wnd[0]/usr").FindByNameEx("COBL-KOSTL", 32).Text = cost_center
            elif aac == "F":
                session.findById("wnd[0]/usr").FindByNameEx("COBL-AUFNR", 32).Text = internal_order
            elif aac == "P":
                session.findById("wnd[0]/usr").FindByNameEx("COBL-PS_POSID", 32).Text = wbs
            elif aac == "A":
                session.findById("wnd[0]/usr").FindByNameEx("COBL-ANLN1", 32).Text = asset
                session.findById("wnd[0]/usr").FindByNameEx("COBL-ANLN2", 32).Text = "0"
            # msg = press()_enter_key()
        if cost_split_count == 1 and tax_code:
            session.findById("wnd[0]/usr").FindByNameEx("TABIDT7", 91).Select()
            session.findById("wnd[0]/usr").FindByNameEx("MEPO1317-MWSKZ", 31).Text = tax_code
            session.findById("wnd[0]").sendVKey(0)
        last_cost_split_item = False
        scroll_next = False
        if aac == "K" and percentage != "100":
            if idx < len(po_lines) - 1:
                if po_lines[idx+1][0] != po_item:
                    last_cost_split_item = True
            else:
                last_cost_split_item = True
            if last_cost_split_item:
                scroll_next = True
        else:
            scroll_next = True
        if scroll_next:
            session.findById("wnd[0]/usr").FindByNameEx("SAPLMEGUITC_1211", 80).verticalScrollbar.Position = Position + 1
            Position = session.findById("wnd[0]/usr").FindByNameEx("SAPLMEGUITC_1211", 80).verticalScrollbar.Position
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except:
        pass
    messagetype = session.findById("wnd[0]/sbar").messagetype
    if messagetype == "W":
        session.findById("wnd[0]").sendVKey(0)
    try:
        session.findById("wnd[1]/usr/btnSPOP-VAROPTION1")
        session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
    except:
        pass
    result = session.findById("wnd[0]/sbar").Text
    if "under the number" in result:
        sap_PO = result.split()[-1]
        result = result + " for po:" + sap_PO
    else:
        result = "failed creating new po"
    print(result)  

def connect_db():
	db_server = config.get('DEFAULT','db_server')
	db_name = config.get('DEFAULT','db_name')
	conn= pymssql.connect(server=db_server,database=db_name)
	cursor = conn.cursor()
	return conn, cursor

def get_from_db(cursor, gr):	
	sql = """select c_PONo,c_ItemNo,c_ReceivingQuantity from app_fd_P02_SelectedGREvent
		where (c_InputStatus is Null or c_InputStatus!='OK') and c_GRProcessIDNo = '%s'""" % gr
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
    headepo_line = [('Content-Type', 'text/plain')]
    start_response(status, headepo_line)
    if environ['REQUEST_METHOD'] == 'POST':
        request_body_size = int(environ.get('CONTENT_LENGTH', 0))
        request_body = environ['wsgi.input'].read(request_body_size)
        #d = papo_linee_qs(request_body)  # turns the qs to a dict
        return 'From POST: %s' % ''.join('%s: %s' % (k, v) for k, v in d.iteritems())
    else:  # GET
        #d = papo_linee_qs(environ['QUERY_STRING'])  # turns the qs to a dict
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
	#port = config.get('DEFAULT','web_port')
	#httpd = make_server('', int(port), simple_app)
	#print "Serving on port %s..." % port
	#httpd.serve_forever()
	session = start_sap()
	if session:
		execute_transaction(session)
		close_sap(session)
	
if __name__ == "__main__":
  main()
