import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

fromaddr = "email@address.com"
toaddr = "email@address.com"
 
msg = MIMEMultipart()
 
msg['From'] = fromaddr
msg['To'] = toaddr
msg['Subject'] = "Meraki Uplink Check"
 
body = "The Meraki devices in the attached report have WAN uplinks that need to be investigated.\n\nOther text can be added here."
 
msg.attach(MIMEText(body, 'plain'))
 
filename = "Meraki Uplink Alert List.xlsx"
attachment = open("Meraki Uplink Alert List.xlsx", "rb")
 
part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
 
msg.attach(part)
 
server = smtplib.SMTP('10.1.1.1', 25) # server IP here
text = msg.as_string()
server.sendmail(fromaddr, toaddr, text)
server.quit()
print("\nEmail sent.  It will arrive in a minute or two.")