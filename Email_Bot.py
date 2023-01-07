
# Libraries used
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import imaplib
import pandas as pd
import time

# Read the name, email, weight, price in Excel
planilha = pd.read_excel('C:/Users/vinic/Desktop/EMAIL PAYMENT.xlsx')
# Assign the values to their respective variable
name = planilha['Name']
email = planilha['Email']
weight = planilha['Weight(KG)']
price = planilha['Price'][0]

host = "(Your host)"
port = 465 # Port of gmail
login = "(Your login)"
password = "(Your Password)"


# Connect to the server
server = smtplib.SMTP_SSL(host, port) # port=465
# Log into the server
server.login(login, password)

# Create a loop
for i in range(len(name)):
    # Body of the message in html
    text = f'''
        <html>
      <body>
        <font color="#000000"><h3>Olá!<br></h3>
        <h3>Segue os dados para depósito ref. pedido de chumbo:<br></h3></font>
        <h2><font color="#0000FF">{weight[i]} kg x {price} = R$ {round(weight[i] * price, 2)}</font></h2> 
        <br>
        <h2><b><font color="#FF0000"><font face="arial">Banco do Brasil</font></font></b></h2>
        <h2><b><font color="#000000">
        Agência #####<br>
        Conta corrente #####<br>
        PIX: #####<br>
        </b></h2></font><br>
        <font color="#000000"><h3>(Name of the company)<br></h3></font>
        --<br>
        <h3><font color="#000000">
        À disposição para esclarecer quaisquer dúvidas.<br>
        </h3></font>
      </body>
    </html>
        '''

    # Structure of the email
    email_msg = MIMEMultipart()
    email_msg['From'] = login
    email_msg['To'] = email[i].strip()
    email_msg['Subject'] = f"{name[i]} - PAYMENT (Name of the Company)"
    email_msg.attach(MIMEText(text, 'html')) # Sending the message in html
    server.send_message(email_msg)

    # Saving in sent box
    imap = imaplib.IMAP4_SSL(host, 993) # Connecting to IMAP server
    imap.login(login, password) # Log into IMAP server
    # Save the email in sent box
    text = email_msg.as_string()
    imap.append('INBOX.Sent', '\\Seen', imaplib.Time2Internaldate(time.time()), text.encode('utf8'))

    # Print for who the email were sent
    print(f'Email sent to {name[i]}')

# Quit of the two servers
imap.logout()
server.quit()
