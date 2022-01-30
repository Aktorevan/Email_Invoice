    # -*- coding: utf-8 -*-
"""
Created on Sun Oct 25 11:02:23 2020

@author: Ninja Xpress
"""

import pandas as pd
import smtplib
import os
import glob
import getpass
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

dir = os.getcwd()

#Data source
data = pd.read_excel("data source1.xlsx".format(dir))

#Email pengirim dan penerima
email_sender = "****@ninja*.codf"
sender_passwd = "****"
email_receiver = data["email"]
recipients = ['****@ninja*.codf']


#Nama shipper dan periode
invoice = data ["invoice number"]
shipper = data["shipper name"]
period = data["period"]

#dir = os.getcwd()

#Isi dari email
body = """
Dear Mitra Ninja,

Berikut Kami lampirkan bukti potong pajak PPh 21 periode tahun 2020 .

Mohon abaikan bukti potong periode tahun 2020 yang sudah kami kirim sebelumnya.
Demikian informasi yang dapat Kami sampaikan, jika Anda membutuhkan bantuan lebih lanjut silahkan balas email ini.


Sincerely,
NinjaXpress
"""

i = 0
while i < data.shape[0]:
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        
        #Untuk mengisi form subject email, pengirim (from), penerima (to)
        msg = MIMEMultipart()
        msg["Subject"] = "Bukti Potong PPh 21 Periode {} - {} ".format(str(period[i]), shipper[i])
        msg["From"] = "Finance - Mitra"
        msg["To"] = email_receiver[i]
        msg["Cc"] = ", ".join(recipients)
        
        #Melampirkan body text
        msg.attach(MIMEText(body.format(shipper[i], str(period[i]), "plain")))
        
        #Mengganti direktori berdasarkan invoice number dan owner name pada data source.xlsx
        os.chdir(str(dir)+"/"+"all bukpot")
        
        #Mengupload semua file yang ada pada folder tersebut
        for files in glob.glob(str(invoice[i])+"*"):
           
            filename = os.path.basename(files)

            #Membuka file dalam binary mode
            with open (filename, "rb") as attachment: 					
                
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
        	
        	#Encode to base64
            encoders.encode_base64(part)								
        
            part.add_header("Content-Disposition", f"attachment; filename = {filename}")
            
            #Melampirkan dokumen
            msg.attach(part)											
        
        #Konversi body text dalam mode string
        text = msg.as_string()											

        #Login ke server
        server.login(email_sender, sender_passwd)

        #Menggabungkan email penerima dan BCC
        receiver = [email_receiver[i]] + recipients
        
        #Mulai mengirim
        server.sendmail(email_sender, receiver, text)
        print("Email terkirim :", email_receiver[i])
    
    i += 1