#!

import openpyxl as xl
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib


#dat=input ("Ingrese Año y Mes: (AAAA/MM): ")
dat="2023/01"
def datesod (dia):
    fecha=dat+"/"+str(dia)
# convert yyyy-mm-dd string to date object
    dt_object = datetime.strptime(fecha, "%Y/%m/%d").date()
    x=dt_object.weekday()
    if x==5:
        print("*** Sábado "+str(dt_object.day)+" ***")
    if x==6:
        print ("*** Domingo "+str(dt_object.day)    +" ***")
    if x==5 or x==6:
        return True
        
def finde(tabla,hoy):
    control=["Ctrl 0/7", "Ctrl 7/16" ,"Ctrl 16/24" , "Vol 0/7", "Vol 7/16", "Vol 16/24"]
    
    for i in range (int(hoy.day)+1,31):
        if datesod(tabla[i][0]):
            trab=[]
            for j in range(1,len(tabla[i])):
                if not tabla[i][j]==None:
                    if not tabla[i][j]=="Franco" :
                        print (tabla[0][j]," : ",tabla[i][j]) 
                        trab.append(tabla[i][j])
            for item in control:
                if not item in trab:
                    print (item, " SIN CUBRIR")
            print("-"*20) 
            
def francos(tabla):
    for j in range (1,len(tabla[0])):
        cont=0
        for i in range(0,31):
            if not tabla[i][j]==None:
                if tabla[i][j]=="Franco":
                    cont+=1
                if i==0:
                    print (tabla[i][j])
                    print ("\n")
                else:
                    print (i,"----",tabla[i][j]) 
        if cont<6:
            print(6-cont," franco a trabajar")
        print ("="*15)   
        print("\n")
               
def send_mail(msg, reciever):
   date=datetime.today()
   email= MIMEMultipart()
   email["From"]=""
   email["To"]=reciever
   msg["Subject"]="Francos ",date.year,"/",date.month
   password=""
   
        
wb=xl.Workbook()
wb=xl.load_workbook("hoja_de_prueba.xlsx")
hoja = wb.active
max_row=31
max_col=16
tabla=[]
for i in range(1,max_row+1):
    row=[]
    for j in range(1,max_col+1):
        cell=hoja.cell(row=i,column=j)
        row.append(cell.value)
    tabla.append(row)


hoy=datetime.today() 
  
#finde(tabla,hoy)

francos(tabla)

#msg="Mensaje de Prueba<br>Prueba de salto de linea"
#send_mail(msg)
print ("Fin")

