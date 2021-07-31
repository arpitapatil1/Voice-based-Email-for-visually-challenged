#!/usr/bin/env python
# coding: utf-8

# In[1]:


#importing required packages
import speech_recognition as sr #speech to text
import yagmail #gmail/SMPT client to send mails
import smtplib#for sending mails using SMTP protocol.SMTP procotolis used so as to communicate with gmail server
from gtts import gTTS#google text to speech. gTTS is module
import os 
import email
import imaplib #access mails over imap protocol on client side
from playsound import playsound  
import random
from email.header import decode_header #for decoding mail in proper format


# In[2]:


#username and password of user (can be already put in the code rather than pronouncing the password)
username="nsp262003@gmail.com"
password="nsp262003" 

session = smtplib.SMTP('smtp.gmail.com',587) #Initiate connection to the server..#host and port area
session.ehlo()  #Hostname to send for this command defaults to the FQDN of the local host
session.starttls() #Start encrypting everything you're sending to the server
session.login(username, password) #Log into the server by sending them our username and password


# In[3]:


#method to clean the string so that text_to_speech() method can take a clean string as argument
def cleanString(incomingString):
    newstring = incomingString
    newstring = newstring.replace("!","")
    newstring = newstring.replace("@","")
    newstring = newstring.replace("#","")
    newstring = newstring.replace("$","")
    newstring = newstring.replace("%","")
    newstring = newstring.replace("^","")
    newstring = newstring.replace("&","and")
    newstring = newstring.replace("*","")
    newstring = newstring.replace("(","")
    newstring = newstring.replace(")","")
    newstring = newstring.replace("+","")
    newstring = newstring.replace("=","")
    newstring = newstring.replace("?","")
    newstring = newstring.replace("\'","")
    newstring = newstring.replace("\"","")
    newstring = newstring.replace("{","")
    newstring = newstring.replace("}","")
    newstring = newstring.replace("[","")
    newstring = newstring.replace("]","")
    newstring = newstring.replace("<","")
    newstring = newstring.replace(">","")
    newstring = newstring.replace("~","")
    newstring = newstring.replace("`","")
    newstring = newstring.replace(":","")
    newstring = newstring.replace(";","")
    newstring = newstring.replace("|","")
    newstring = newstring.replace("\\","")
    newstring = newstring.replace("/","")
    newstring = newstring.replace(".","")
    return newstring


# In[4]:


#convert string to speech
def text_to_speech(sentence):
    sentence=cleanString(sentence)
    
    obj = gTTS(text=sentence, lang='en')
    r1 = random.randint(1,10000000)
    r2 = random.randint(1,10000000)
    if(len(sentence)>20):
        sentence=sentence[:20]
    filename = str(r2)+sentence+str(r1) +".mp3"
    
    obj.save(filename) 
    playsound(filename)
    os.remove(filename)


# In[5]:


#get the audio from the user
def speech_to_text():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print('Clearing background noise..')
        r.adjust_for_ambient_noise(source,duration=1)
        print("Say something!")
        audio = r.listen(source)
        print('Done recording..!')

    try:
        print("You said: " + r.recognize_google(audio,language='en-GB'))
    except sr.UnknownValueError:
        print("Google Speech Recognition could not understand audio")
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))
    return str(r.recognize_google(audio))


# In[6]:


#send email to specified receiver
def send_message():
    text_to_speech('say the E-Mail address of the receiver')
    receiver_email=speech_to_text()
    words=receiver_email.split()
    modified_mail=str()
    for word in words:
        if word == 'underscore':
            modified_mail=modified_mail+'_'
        elif word == 'dot':
            modified_mail=modified_mail+'.'
        else:
            modified_mail=modified_mail+word
    modified_mail=modified_mail.lower()
    print ("Reciever email is: "+modified_mail)
    text_to_speech('Say the subject')
    subject=speech_to_text()
    text_to_speech('Say the message')
    message=speech_to_text()
    sender=yagmail.SMTP(username,password)
    sender.send(to=modified_mail,subject=subject,contents=message)
    print("Message sent!")


# In[7]:


#count the number of unseen mails
def get_unseen_no():
    mail = imaplib.IMAP4_SSL('imap.gmail.com',993) #this is host and port area.... ssl security
    mail.login(username,password)  #login
    mail.select('Inbox') 
    stat,unseen = mail.search(None, 'UnSeen') 
    unseen=str(unseen) #convert list to string
    str1=unseen.strip("[,b,',")
    lst=str1.split(" ")
    text_to_speech("Number of UnSeen mails is:"+str(len(lst)))
    print ("Number of UnSeen mails :"+str(len(lst)))


# In[8]:


#read the most recent mail in the inbox
def read_recent():
    # create an IMAP4 class with SSL 
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    # authenticate
    imap.login(username, password)
    status, messages = imap.select("INBOX") 
    N = 1 # number of top emails to fetch ..change N to read more than 1 recent mails
    messages = int(messages[0]) # total number of emails
    for i in range(messages, messages-N, -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)") #returns tuple
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                str1="Subject is "+subject              
                str2=""+From
                li=str2.split("<")
                str1=str1+" "+"From "+li[0]
                print(str1)
                type(str1)
                text_to_speech(str1)
                if msg.is_multipart(): #if there are many parts in the mail 
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition")) #this is either inline opening in browser or attachment
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            x=body.splitlines()
                            body=""
                            body="".join(x)
                            print(body)
                            if(body != ""):
                                text_to_speech(body)
                            else:
                                print("No body Present")
                                text_to_speech("No Body present")
                       
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        x=body.splitlines()
                        body=""
                        body="".join(x)
                        print(body)
                        if(body != ""):
                            text_to_speech(body)
                        else:
                            print("No body Present")
                            text_to_speech("No Body present")
                   
                print("="*100)


    imap.close()
    imap.logout()            


# In[9]:


def search_particular():
    #read particular mail by giving the subject
    #text_to_speech('Say the name of sender')
    #sender=speech_to_text()
    #text_to_speech('Say the subject to be searched')
    #subj=speech_to_text()

    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(username,password)
    mail.select("inbox") # connect to inbox.
    text_to_speech('say the E-Mail address for search')
    receiver_email=speech_to_text()
    words=receiver_email.split()
    modified_mail=str()
    for word in words:
        if word == 'underscore':
            modified_mail=modified_mail+'_'
        elif word == 'dot':
            modified_mail=modified_mail+'.'
        else:
            modified_mail=modified_mail+word
    modified_mail=modified_mail.lower()
       
    print ("Search email is: "+modified_mail)
    text_to_speech('Say the subject')
    subject=speech_to_text()
   
    result, data = mail.search(None, '(FROM "'+modified_mail+'" SUBJECT "'+subject+'")' )
    
    ids = data[0] # data is a list.
    id_list = ids.split() # ids is a space separated string
    latest_email_id = id_list[-1] # get the latest email

    result, data = mail.fetch(latest_email_id, "(RFC822)") # fetch the email body (RFC822) for the given ID
    for response in data:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                data = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(data["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # decode email sender
                From, encoding = decode_header(data.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                print("Subject is", subject)
                str1="Subject is "+subject
                print("From ", From)
                str2=""+From
                li=str2.split("<")
                str1=str1+" "+"From "+li[0]
                print(str1)
                type(str1)
                text_to_speech(str1)
                if data.is_multipart():
                    # iterate over email parts
                    for part in data.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            x=body.splitlines()
                            body=""
                            body="".join(x)
                            print(body)
                            if(body != ""):
                                text_to_speech(body)
                            else:
                                print("No body Present")
                                text_to_speech("No Body present")
                       
                else:
                    # extract content type of email
                    content_type = data.get_content_type()
                    # get the email body
                    body = data.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        x=body.splitlines()
                        body=""
                        body="".join(x)
                        print(body)
                        if(body != ""):
                            text_to_speech(body)
                        else:
                            print("No body Present")
                            text_to_speech("No Body present")
                    
                print("="*100)


# In[21]:


while True :
    text_to_speech('say 1 1 to send a message !')
    text_to_speech('say 1 2 to read recent mail !')
    text_to_speech('say 1 3 to check number of unseen mails in inbox !')
    text_to_speech('say 1 4 to search and read particular mail !')
    text_to_speech('say 1 5 to close the application !')
    first_response=speech_to_text()
    if first_response == '11' :
        send_message()
    elif first_response == '12':
        read_recent() 
    elif first_response == '13'  :
        get_unseen_no()
    elif first_response == '14':
        search_particular()
    elif first_response == '15':
        break
    else:
        text_to_speech('Sorry you were not clear with your vocals !')
        continue


# In[ ]:





# In[ ]:




