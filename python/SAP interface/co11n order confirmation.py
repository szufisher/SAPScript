import sys, win32com.client
import logging
import pypyodbc
from datetime import datetime
import time
import smtplib
from email.mime.text import MIMEText
import schedule
import ConfigParser
from utils import handle_warning, logon_sap,start_sap,close_sap, connect_db, close_db,get_configer,get_logger, timestamp, send_email

cf =get_configer('co11n.conf')
logger = get_logger('co11n.log')

def execute_transaction(session, cursor, order_lines):
    for idx, order_line in enumerate(order_lines):
        ConfirmMessage, ConfirmStatus = '',''        
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nco11n"
        session.FindById("wnd[0]").sendVKey(0)        
        session.findById("wnd[0]/usr").FindByNameEx("AFRUD-AUFNR",32).Text = order_line[0]
        session.findById("wnd[0]/usr").FindByNameEx("AFRUD-VORNR",32).Text = order_line[1]
        session.FindById("wnd[0]").sendVKey(0)
        if session.ActiveWindow.Name =="wnd[1]": #already confirmed or skip preceeding step
            session.FindById("wnd[0]").sendVKey(0)
            #session.findById("wnd[1]/usr/btnOPTION2").press()
        handle_warning(session)
        operation_des = session.findById("wnd[0]/usr").FindByNameEx("AFVGM-LTXA1",31).Text
        status_text = session.FindById("wnd[0]/sbar").Text
        if session.FindById("wnd[0]/sbar").MessageType == "E":
            ConfirmMessage = status_text
            ConfirmStatus = "Confirm"
        else:            
            qPn = session.findById("wnd[0]/usr").FindByNameEx("CAUFVD-MATNR",31).Text            
            if qPn not in order_line[2]:               
                ConfirmMessage = "Material number in EDS does not match SAP"
                ConfirmStatus = "Pending"        
            order_Qty=session.findById("wnd[0]/usr").FindByNameEx("CORUF-SMENG",31).text
            confirmed_Qty=session.findById("wnd[0]/usr").FindByNameEx("AFVGD-LMNGA",31).text
            if order_Qty == confirmed_Qty:    # ensure there is open to be confirmed qty
                ConfirmMessage = "Already Confirmed"
                ConfirmStatus = "Closed"            
        if ConfirmMessage and ConfirmStatus:
            update_to_db(cursor, order_line[0],order_line[1],order_line[3],ConfirmStatus, ConfirmMessage, operation_des)
            continue            
        if float(order_Qty) > float(confirmed_Qty) + 1:
            session.findById("wnd[0]/usr").FindByNameEx("AFRUD-AUERU",34).key = ""            
        session.findById("wnd[0]/usr").FindByNameEx("AFRUD-AUSOR",42).Selected = 0
        session.findById("wnd[0]/usr").FindByNameEx("AFRUD-LMNGA",31).Text = "1"         
        session.FindById("wnd[0]").sendVKey(0)
        handle_warning(session)
        if session.ActiveWindow.Name =="wnd[1]": #Activities are recalculated due to quantity change        
            session.ActiveWindow.Close()                        
        session.FindById("wnd[0]/tbar[0]/btn[11]").press()
        handle_warning(session)
        status_text = session.FindById("wnd[0]/sbar").Text
        ConfirmMessage = status_text
        ConfirmStatus = "Closed"  
        if "Error in goods movements" in status_text:
            session.findById("wnd[0]/usr/btnE_DETAIL").press()
            error = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").GetCellValue(0,'CMF_TEXT')
            if "eriod" in error:    # Period 012/2019 is not open for account type S
                send_email('skipped due to period end closing', "co11n@xx.com")
                break
            session.FindById("wnd[0]/tbar[0]/btn[3]").press() #back    
            session.FindById("wnd[0]/tbar[0]/btn[3]").press() #back
            # click No button to Confirmation saved(goods movement 0, failed 3)
            session.FindById("wnd[1]/usr/btnSPOP-OPTION2").press()
        elif "saved" not in status_text: 
            ConfirmStatus = "IT support Needed"
        else:                  
            if session.FindById("wnd[0]/sbar").MessageType == "E":
                ConfirmStatus = "Pending"
        
        update_to_db(cursor, order_line[0],order_line[1],order_line[3], ConfirmStatus, ConfirmMessage, operation_des)
        if "Goods movements" in ConfirmMessage:
            time.sleep(2)    #wait till background process finish, otherwise next confirmation will be blocked by own account
        
def get_from_db(cursor):
    # dateadd(day,datediff(day,10,GETDATE()),0)
    sql = """SELECT distinct A.WO, A.OpStep, A.PN,A.SN, C.last_step FROM IdsTlg53_ApsRoutingStart AS A 
        inner join Mis03Routing as D on A.PN = D.Pn            
        inner join (select r.Pn, max(r.OpStep) as last_step from Mis03Routing as r group by r.Pn) as C on A.PN = C.Pn
        left join IdsTlg52_ApsPoSnSta AS E on A.WO = E.PO and A.SN = E.SN 
        left join Mis12RoutingSta AS B ON (A.OpStep = B.OpStep) AND (A.SN = B.SN) AND (A.WO = B.WO) and  (B.SapSta = 'Closed')             
        where A.Status ='Complete' and len(A.SN) = 4 and A.FinishTime>='2019-08-14' and D.SapConfirm = 1 and 
        ((A.OpStep <> C.last_step) or (A.OpStep = C.last_step and E.PO is not null))
        and (B.WO is null ) order by A.OpStep"""    
    cursor.execute(sql)    
    result = cursor.fetchall()
    return result
    
def update_to_db(cursor, wo, operation,sn, status, message, operation_des):
    print('%s,%s,%s,%s,%s' %(wo,sn, operation, status, message))
    logger.info('%s,%s,%s,%s,%s' %(wo,sn, operation, status, message))
    cursor.execute("""insert into Mis12RoutingSta(WO, SN, OpStep, SapSta, SapStaRemark, SapStaTime, OpDes )
                    values(?, ?, ?, ?, ?, ?, ?)""", (wo,sn,operation, status, message, datetime.now(), operation_des))
    return cursor.rowcount

        
def job():
    try:
        print('%s started running the job...' % datetime.now())    
        short_cut_file =cf.get('saplogon','short_cut_file')
        popup_win_title=cf.get('saplogon','popup_win_title')
        pin =cf.get('saplogon','pin')
        wait_sec =cf.get('saplogon','wait_sec')
        logon_sap(short_cut_file, popup_win_title, pin, wait_sec)
        time.sleep(2)                
        j = 30
        while j > 0:
            session = start_sap()
            if session:
                break
            else:
                j-= 1
                time.sleep(1)
                
        if session:            
            conn, cursor = connect_db(cf.get('db','ip'), cf.get('db','db'),cf.get('db','uid'),cf.get('db','pwd'))
            order_lines = get_from_db(cursor)
            for i in order_lines:
                print(i)
            #order_lines =[['22232516','10','11255594', '']]
            execute_transaction(session, cursor, order_lines)
            close_db(conn)            
            close_sap(session)            
            send_email('CO11N processed %s records' % len(order_lines), "co11n@xx.com")
        else:
            send_email("Failed logon SAP", 'co11n@xx.com')
        print('%s finished running the job...' % datetime.now())                
    except Exception, e:
        send_email("order confirmation run with error %s" % str(e), 'co11n@xxx.com')
        raise
    
def main():
    run_at=cf.get('schedule','RunAt')
    runat = run_at.split(';')
    for r in runat:
        schedule.every().day.at(r).do(job)
    print('%s waiting for pending job at %s' %(datetime.now(),runat))    
    while True:
        schedule.run_pending()
        time.sleep(1)

    
if __name__ == "__main__":
  main()
