import win32com.client as win32
import pandas as pd
import os

def Excel_pull(Excel_name, Column_name,cc_Column):
    global Tolist
    global Cclist
    df = pd.read_excel(Excel_name)
    list = df[Column_name].tolist()
    cclist= df[cc_Column].tolist()
    Output=[]
    Cc_Output=[]
    for i in list:
        try:
            Output.append(i+";")
            Tolist=' '.join(Output)
        except:
            continue
    for i in cclist:
        try:
            Cc_Output.append(i+";")
            Cclist=' '.join(Cc_Output)
        except:
            continue
    return Tolist,Cclist

def Word_Extract(Word_Name):
    directory=os.getcwd()
    difile=directory+"\\"+Word_Name
    word = win32.Dispatch("Word.Application")
    doc = word.Documents.Open(difile)
    doc.Content.Copy()
    doc.Close()

def EzMail(text, subject, recipient, cc_recipient):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    #mail.HtmlBody = text
    mail.cc=cc_recipient
    mail.GetInspector.WordEditor.Range(Start=0, End=0).Paste()
    mail.Display(True)

if __name__=='__main__':
    #PLEASE CLOSE THE EXCEL BEFORE RUNNING
    print("CLOSE EXCEL AND OPEN OUTLOOK!")
    Word_Extract("Auto.docx")
    Excel_pull('Email Sender.xlsx','To_Email','CC_Email')
    EzMail('<h2>Algo by Chris<h2>',"COP021",Tolist,Cclist)