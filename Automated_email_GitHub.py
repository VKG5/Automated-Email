#!/usr/bin/env python
# coding: utf-8

# ## A Program to send emails to everyone in a file

# In[ ]:


# For making the normal program work
import pandas as pd
import smtplib

# For attachments
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 

# The excel sheet with the emails
file = pd.read_excel('***.xlsx',)


# ### Generating the required list

# In[ ]:


file.head()
email_list = file["Column Name"].values


# In[ ]:


# Adding the necessary @gmail, @outlook, suffix.
suffix = "@gmail.com"

final_list = []
for f in email_list:
    if suffix not in f:
        final_list.append(f.strip() + suffix)
    else:
        final_list.append(f)


# ### Sorting the list

# In[ ]:


juniors = []
curr = []
seniors = []

# If you have sub-class emails in the excel sheet, change the number fo conditions as per your convenience
for email in final_list:
    if '***' in email:
        juniors.append(email)
    elif '***' in email:
        curr.append(email)
    else:
        seniors.append(email)


# ### Logging in to your account

# In[ ]:


# If you want to dens this from some other email service, just look up the equivalent port on google :D
server = smtplib.SMTP('smtp.gmail.com', 587)
print("Done")
server.ehlo()
server.starttls()

try:
    server.login('your_email_address@gamil.com','Enter your password here')
    print("Login Successful!")

except:
    print("Wrong Credentials!")


# ### Generating a message

# Simple body and subject of the email

# In[ ]:


msg = '''The message you want in the email'''

subject = 'Subject of the email'

body = 'Subject: {}\n\n{}'.format(subject,msg)


# The **attachment** in the email

# In[ ]:


msg = MIMEMultipart()

msg.attach(MIMEText(body, 'plain'))

# Name of the file that is to be attached, opening the file to be attached
file_to_be_attached = '***.(extension of the file to be attached)'
attachment = open('***.(extension of the file to be attached)', 'rb')

p = MIMEBase('application', 'octet-stream') 

p.set_payload((attachment).read()) 

encoders.encode_base64(p) 

p.add_header('Content-Disposition', "attachment; filename= %s" %file_to_be_attached)

msg.attach(p) 


# ### Test ID(s)

# In[ ]:


# Test mail
email = ['***@gmail.com']

# To check number of emails 
print(len(curr))

final_mail = msg.as_string()

count=0
for e in email:
    server.sendmail('your_email_address@gmail.com', e, final_mail)
    print("Sent")
    count+=1

# To check if the count is equal in both
print(count)
server.quit()


# In[ ]:




