#Credit
# https://www.freecodecamp.org/news/send-emails-using-code-4fcea9df63f/
# https://blog.mailtrap.io/sending-emails-in-python-tutorial-with-code-examples/

import os 		#File Directory
import smtplib	#SMTP Server

from string import Template

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

emails = [] #List of emails to message, extracted from pdf files

#####################################################################
#CHANGE THESE VALUES
MY_ADDRESS = 'email@domain.ac.za'				#Sender Email
PASSWORD = 'Password123'						#Sender Password
UNIVERSAL_DOMAIN = '@domain.ac.za' 				#Receiver emails
FILE_LOCATION = 'PDF_Files/'				
MESSAGE_TEMPLATE_LOCATION = 'message.txt'
MESSAGE_SUBJECT = "This is a test"
#####################################################################

def get_emails_from_files(directory):
	files = os.listdir(directory)
	for item in files:
		if item.endswith(".pdf"):
			emails.append(os.path.splitext(item)[0]) #Remove extension and add to list
	return emails

def read_message(filename): #Create email message form txt template
	with open(filename, 'r', encoding='utf-8') as template_file:
		template_file_content = template_file.read()
	return Template(template_file_content)


def main():
	emails = get_emails_from_files(FILE_LOCATION)				#Get Emails
	message_template = read_message(MESSAGE_TEMPLATE_LOCATION)	#Create Message

	#Set up the SMTP server
	s = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
	s.starttls()
	s.login(MY_ADDRESS,PASSWORD)

	for name in emails:
		msg = MIMEMultipart() 		#create a message

		#message = message_template.substitute(PERSON_NAME=name.title())		#Sustitute Name Variable from txt file ------${PERSON_NAME}

		#print(message) 				#For testing purposes
		print('Sending to '+name)		#For Progress Report

		email = name + UNIVERSAL_DOMAIN		#Create email address from file name

		#Setup for the email parameters
		msg['From']=MY_ADDRESS
		msg['To']=email
		msg['Subject']=MESSAGE_SUBJECT

		#Add Pdf File
		filename = FILE_LOCATION+name+'.pdf'
		with open(filename, "rb") as attachment:
			part = MIMEBase("application","octet-stream")
			part.set_payload(attachment.read())

		encoders.encode_base64(part)
		part.add_header('Content-Disposition','attachment', filename=name+'.pdf')
		msg.attach(part)


		#Add txt message
		msg.attach(MIMEText(message, 'plain'))

		#s.sendmail(msg)
		s.sendmail(MY_ADDRESS, [email], msg.as_string())
		print('Sent to ' + name)
		del msg

	#Terminate SMTP session
	s.quit()


if __name__ == '__main__':
	response = input("Did you check the variables? (Y/N) ")
	if response == "Y" or response == "y":
		main()
