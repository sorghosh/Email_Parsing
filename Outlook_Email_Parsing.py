import win32com.client
import os
from sqlite3 import dbapi2 as sqlite
import re

class email_parser:
    def __init__(self,db):
        self.con = sqlite.connect(db)
    
    def __del__(self):
        self.con.close()
    
    def dbcommit(self):
        self.con.commit()
        
    def createtable(self):
        self.con.execute('create table Email (Email_Name,SenderName,Sender_Email_Address,SentOn,Tos,CC,BCC,Subject)')
        self.con.execute('create table Word (Email_Id,Line_No, Text)')
        self.dbcommit()
    
    def insertion(self,tablename , dataset):       
        self.con.execute("insert into" +"  " + tablename +"  "+ "values ("+('?,' * len(dataset))[:-1]+")", dataset)
        rowid = self.con.execute("select max(rowid) from"+' '+ tablename)
        rowid = rowid.fetchone()[0]
        return rowid

    def check_email_exists(self,emailname):
        cur = self.con.execute("select count(*) from email where email_name = ?", (emailname,))
        rec = cur.fetchone()
        return rec[0]
            
    def gettext(self,text,email_id):
        word = text[0].split("\r\n")
        spliter = re.compile("\\W*")
        dataset = []
        splitword = []
        for w in word:
            if len(w) > 0 :
                splitword_temp = [str(i).lower() for i in spliter.split(w) if i != '']
                splitword.append(splitword_temp)
        for line_id in range(0,len(splitword)):
            for w in splitword[line_id]:
                dataset.append(email_id)
                dataset.append(line_id)
                dataset.append(w)
                self.insertion("word",dataset)
                dataset = []
    
    def read_email(self,dirs):   
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        email_temp = [] 
        email_body = []
        for root, dr, files in os.walk(dirs):
            for f in files:
                filename = os.path.join(root,f)
                msg = outlook.OpenSharedItem(filename)
                email_temp.append(f)
                email_temp.append(msg.SenderName)
                email_temp.append(msg.SenderEmailAddress)
                email_temp.append(msg.SentOn)
                email_temp.append(msg.To)
                email_temp.append(msg.CC)
                email_temp.append(msg.BCC)
                email_temp.append(msg.Subject)
                email_temp = [str(e) for e in email_temp] 
                email_body.append(msg.Body)
                row_count = self.check_email_exists(f)
                if row_count > 0:
                    continue
                email_id = self.insertion("email",email_temp)
                self.gettext(email_body,email_id)
                email_temp = []
                email_body = []
                self.dbcommit()        



dirs    = "Enter the directory name where the email messages are stored"   
dbname  = "Enter the path name where the database is intended to be stored"    
eml     = email_parser(dbname)
eml.read_email(dirs)

