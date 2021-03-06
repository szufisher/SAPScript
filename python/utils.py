import sys, win32com.client
import logging
import pypyodbc
from datetime import datetime
import time, calendar
import smtplib
from email.mime.text import MIMEText
import configparser
import os
from pywinauto import application

def timestamp():
    return calendar.timegm(time.gmtime())

def logon_sap(short_cut_file, popup_win_title,pin, wait_sec):
    i = 3
    while i > 0:
        if i == 2 and start_sap():
            return
        if _logon_sap(short_cut_file, popup_win_title,pin, wait_sec):
            return
        i -= 1    

def close_previous_failed_sap_session():
    app = application.Application()
    try:
        app.connect(title_re='SAP GUI*')
        app.kill()
    except:
        pass

def _logon_sap(short_cut_file, popup_win_title,pin, wait_sec):
    try:
        os.startfile(short_cut_file)
        time.sleep(5)
        session, msg = _get_sap_session(short_cut_file, popup_win_title,pin, wait_sec)
        if session:
            return
            
        app = application.Application()
           
        def get_popup_window():
            try:
                app.connect(title=popup_win_title)
                return True
            except:
                return False
        i = int(wait_sec)
        while i > 0:
            if get_popup_window():
                print(i, 'pin popup window found')
                app[popup_win_title].type_keys(pin)
                app[popup_win_title].type_keys('{ENTER}')
                return True
            else:            
                i -= 1            
                time.sleep(1)
    except:
        pass

def get_configer(filename):
    cf = configparser.ConfigParser()
    cf.read(filename)
    return cf
    
def get_logger(filename):
    logging.basicConfig(filename=filename,format='%(asctime)s %(message)s', level=logging.INFO)
    logger=logging.getLogger(__name__)
    return logger

def start_sap():
    """run as windows scheduled task does not work, maybe due to task only run in background, no way to manipulate the window in dialog mode,
    so dialog mode by long running python program with own schedule defined in config file""" 
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")    
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)        
        session = connection.Children(0)
        return session
    except:        
        print(sys.exc_info()[0])

def get_sap_session(short_cut_file, popup_win_title,pin, wait_sec, count=0):
    close_previous_failed_sap_session()
    session, msg = _get_sap_session(short_cut_file, popup_win_title,pin, wait_sec, count)
    print(count,msg)
    if count==0 and session:   # closing existing active session first, 
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nex"
        session = 0
    while not session and count < 3:
        _logon_sap(short_cut_file, popup_win_title,pin, wait_sec)
        count += 1
        session, msg = _get_sap_session(short_cut_file, popup_win_title,pin, wait_sec, count)
        print('retry',count,msg)    
    return session, msg    
        
def _get_sap_session(short_cut_file, popup_win_title,pin, wait_sec, count=0):    
    try:
        print(count, 'get_sap_session')
        if count >0:
            time.sleep(2)   # after logon, wait 2 seconds
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        print(count, 'sapguiauto')
        application = SapGuiAuto.GetScriptingEngine
        print(count,'application')
        connection = application.Children(0)
        print(count,'connection')
        time.sleep(1)
        session = connection.Children(0)
        print(count,'session')
        return (session, 'ok')
    except:
        print(count, sys.exc_info()[0])
        return (0, '%s %s' %(count, sys.exc_info()[0]))
            
def close_sap(session):
    if not session:
        return
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nex"
    session.FindById("wnd[0]").sendVKey(0)    
    session = None
    connection = None
    application = None
    SapGuiAuto = None

def handle_warning(session):
    for i in xrange(10):
        if session.FindById("wnd[0]/sbar").MessageType == "W":
            session.FindById("wnd[0]").sendVKey(0)
        else:
            break        
        
def connect_db(ip,db, uid=None, pwd=None):
    conn_str = 'Driver={SQL Server};Server=%s;Database=%s' %(ip, db)
    if uid:
        conn_str += ';uid=%s;pwd=%s' %(uid,pwd)
    conn = pypyodbc.connect(conn_str)        
    cursor = conn.cursor()
    return conn, cursor
   
def close_db(conn):
    if not conn:
        return
    conn.commit()
    conn.close()
    
def send_email(subject, sender):
    mailto = 'xxx'    
    msg = MIMEText('auto mail from interface program', 'plain')
    msg['Subject']= subject
    msg['From']   = sender 
    server = smtplib.SMTP('',25)
    #server.starttls()    #no authentication needed, register the IP in smtp server instead
    server.ehlo_or_helo_if_needed()
    try:
        failed = server.sendmail(sender, mailto, msg.as_string())
        server.close()
    except Exception as e:
        print(e)   
