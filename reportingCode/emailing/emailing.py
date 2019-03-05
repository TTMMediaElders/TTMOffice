import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate

#This function is not optimized to send in large quantities.
def send_gmail(send_from, send_to, password, text='', subject='', files=None):
	# assert isinstance(send_to, list)
	msg = MIMEMultipart()
	msg['From'] = send_from
	msg['To'] = send_to
	msg['Date'] = formatdate(localtime=True)
	msg['Subject'] = subject
	msg.attach(MIMEText(text))
	for f in files or []:
		with open(f, "rb") as fil:
			part = MIMEApplication(
				fil.read(),
				Name=basename(f)
			)
		# After the file is closed
		part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
		msg.attach(part)
	smtp = smtplib.SMTP(host='smtp.gmail.com', port=587)
	smtp.starttls()
	smtp.login(send_from, password)
	smtp.sendmail(send_from, send_to, msg.as_string())
	smtp.close()
