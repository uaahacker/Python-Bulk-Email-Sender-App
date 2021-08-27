from tkinter import *
from tkinter import messagebox, filedialog
from pygame import mixer
import speech_recognition
from email.message import EmailMessage
import smtplib
import os
import imghdr
import pandas

###################################################
# Version : 0.0.1
# Author : uaahacker
# Description : Email Sender
# website : https://www.truecoredev.com
# Github : https://github.com/uaahacker
####################################################

# Browse Button Function to browse xlsx file to read emails  init
def browse():
    global final_emails
    path=filedialog.askopenfilename(initialdir='c:/', title='Select Excel File')
    if path == '':
        messagebox.showerror('Error', 'Please Select an Excel File')

    else:
        data=pandas.read_excel(path)
        if 'Email' in data.columns:
            emails = list(data['Email'])
            final_emails = []
            for i in emails:
                if pandas.isnull(i) == False:
                    final_emails.append(i)

            if len(final_emails) == 0:
                messagebox.showerror('Error', 'File does not contain any email addresses !')

            else:
                toEntryField.config(state=NORMAL)
                toEntryField.insert(0, os.path.basename(path))
                toEntryField.config(state='readonly')
                totalLabel.config(text='Total: '+str(len(final_emails)))
                sentLabel.config(text='Sent:')
                leftLabel.config(text='Left:')
                failedLabel.config(text='Failed:')



# bUtton Ch4eck Function  if multiple  radio button is selected then we will enable the state of the Browse button
def button_check():

    if choice.get() == 'multiple':
        browseButton.config(state=NORMAL)
        toEntryField.config(state='readonly')

    if choice.get() == 'single':
        browseButton.config(state=DISABLED)
        toEntryField.config(state=NORMAL)




# check is flase because now the user does'nt selected the file rightnow
check = False


# To Add attachments to The Email
def attachment():
    
    global filename, filetype, filepath, check
    check = True
    filepath=filedialog.askopenfilename(initialdir='c:/', title=' Select File')
    filetype=filepath.split('.')
    filetype=filetype[1]
    filename=os.path.basename(filepath)
    #inserting files or file to text area
    textarea.insert(END, f'\n{filename}\n')


# Sending Email Function
def sendingEmail(toAddress, subject, body):
    f=open('credentials.txt', 'r')

    for i in f:
        credentials=i.split(',')

    message=EmailMessage()
    message['subject'] = subject
    message['to'] = toAddress
    message['from'] = credentials[0]
    message.set_content(body)
   
    # Check if the user has selected a file or not
    if check:

        if filetype == 'png' or filetype == 'jpg' or filetype == 'jpeg':
            f=open(filepath, 'rb')
            file_data = f.read()
            subtype=imghdr.what(filepath)

            message.add_attachment(file_data, maintype='image', subtype=subtype, filename=filename)

        else:
            f=open(filepath, 'rb')
            file_data = f.read()
            message.add_attachment(file_data, maintype-'application', subtype='octet-stream', filename=filename)

    s=smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(credentials[0], credentials[1])
    s.send_message(message)
    x=s.ehlo()
    if x[0] == 250:
        return 'sent'
    else:
        return 'failed'


    





# Send Email Function  
def send_email():
    # to check wether the inputs or empty or not
    if toEntryField.get() == '' or subjectEntryField.get() == '' or textarea.get(1.0, END) == '\n':

        messagebox.showerror('Error', 'All fields are Required', parent=root)

    # check which radio buttons is selected
    else:
        #if single radio button is selected 
        if choice.get() == 'single':
            result=sendingEmail(toEntryField.get(), subjectEntryField.get(), textarea.get(1.0, END))
            if result == 'sent':
                messagebox.showinfo('Success', 'Email is sent successfully')

            if result == 'failed':
                messagebox.showerror('Error', 'Email is not Sent')  

        #if multiple radio button is selected 
        if choice.get() == 'multiple':
            sent=0
            failed=0
            for x in final_emails:
                result=sendingEmail(x, subjectEntryField.get(), textarea.get(1.0, END))
                if result == 'sent':
                    sent+=1
                if result == 'failed':
                    failed+=1


                totalLabel.config(text='')
                sentLabel.config(text='Sent:' + str(sent))
                leftLabel.config(text='Left:' + str(len(final_emails) - (sent + failed)))
                failedLabel.config(text='Failed:' + str(failed))


                totalLabel.update()
                sentLabel.update()
                leftLabel.update()
                failedLabel.update()

        messagebox.showinfo('Success', 'Emails are sent successfully')


# Setting Function that open a new window for login details 
def settings():

    # Clear function for login window  to clear the login info 
    def clear1():
        # to clear the input
        fromEntryField.delete(0, END)
        passwordEntryField.delete(0, END)


    # Save Function for login window to save input in the txt file
    def save():
        if fromEntryField.get() == '' or passwordEntryField.get() == '':
            messagebox.showerror('Error', 'All fields are Required', parent=root1)
        else:
            f=open('credentials.txt', 'w')
            f.write(fromEntryField.get()+ ',' +passwordEntryField.get())
            f.close()
            messagebox.showinfo('Information', 'CREDENTIALS SAVED SUCCESSFULLY', parent=root1)

    root1=Toplevel()
    #title of the new window that appears
    root1.title('Settings')
    # size of the window 
    root1.geometry('650x340+350+90')
    # Background of the window
    root1.config(bg='gray15')

    #title label
    Label(root1, text=' Credentail Settings', image=logoImage, compound=LEFT, font=('goudy old style', 40, 'bold'), fg='gray', bg='gray15').grid(padx=40)

    # Label for input text area from email which we are sending messages
    fromLabelFrame=LabelFrame(root1, text='From (Email Address)', font=('times new roman', 16, 'bold'), bd=3, fg='white', bg='gray15')
    fromLabelFrame.grid(row=1, column=0, pady=20)
    # Entry Fields to enter some text to email 
    fromEntryField=Entry(fromLabelFrame, font=('times new roman', 18, 'bold'), width=30)
    fromEntryField.grid(row=0, column=0)

    # Label for input text area  for password of the email
    passwordLabelFrame=LabelFrame(root1, text='Password', font=('times new roman', 16, 'bold'), bd=3, fg='white', bg='gray15')
    passwordLabelFrame.grid(row=2, column=0, pady=20)
    # Entry Fields to enter some text to Password 
    passwordEntryField=Entry(passwordLabelFrame, font=('times new roman', 18, 'bold'), width=30, show='*')
    passwordEntryField.grid(row=0, column=0)

    # 2 Buttons for  save and clear to save and clear the input
    # SAVE  Button
    Button(root1, text='SAVE', font=('times new roman', 18, 'bold'), cursor='hand2', bg='silver', fg='gray', command=save).place(x=210, y=280)
    # Clear Button
    Button(root1, text='CLEAR', font=('times new roman', 18, 'bold'), cursor='hand2', bg='silver', fg='gray', command=clear1).place(x=340, y=280)

    # Opening credentials.txt file
    f=open('credentials.txt', 'r')
    #spliting txt files and convert into list
    # And itterate the list
    for i in f:
        credentals = i.split(',')

    # Entering email and password field into empty fields
    fromEntryField.insert(0, credentals[0])
    passwordEntryField.insert(0, credentals[1])



    root1.mainloop()

# Button to Speak
#what we ever speck that will be converted into text
def speak():
    mixer.init()
    mixer.music.load('music1.mp3')
    mixer.music.play()
    sr=speech_recognition.Recognizer()
    with speech_recognition.Microphone() as m:
        try:
            sr.adjust_for_ambient_noise(m, duration=0.2)
            audio=sr.listen(m)
            text=sr.recognize_google(audio)
            textarea.insert(END, text + '.')

        except:
            pass




# to exit the window we have used iexit function to close the window
def iexit():
    result=messagebox.askyesno('Notification','Do you want to exit?')
    if result:
        root.destroy()
    else:
        pass


# to clear the text in the text field we have used clear function to clear all the input
def clear():
    toEntryField.delete(0, END)
    subjectEntryField.delete(0, END)
    textarea.delete(1.0, END)







root=Tk()
root.title('Email Sender App')
root.geometry('780x620+100+50')
root.resizable(0,0)
root.config(bg='gray15')

# Top logo and Setting button and Images
titleFrame=Frame(root,bg='gray15')
titleFrame.grid(row=0,column=0)
logoImage=PhotoImage(file='email.png')
titleLabel=Label(titleFrame, text='   Email Sender', image=logoImage, compound=LEFT, font=('Goudy Old Style', 28, 'bold'), bg='gray15', fg='white')
titleLabel.grid(row=0, column=0)

# Setting Button to open a new window where you put your crediantials
settingImage=PhotoImage(file='setting.png')
Button(titleFrame,image=settingImage, bd=0, bg='gray15', cursor='hand2', activebackground='gray15', command=settings).grid(row=0, column=1, padx=20)

#Radio Button Area choose frame
chooseFrame=Frame(root, bg='gray15')
chooseFrame.grid(row=1, column=0, pady=10)
choice=StringVar()
# Radio button 1 single
singleRadioButton=Radiobutton(chooseFrame, text='Single   ', font=('times new roman', 25, 'bold'), variable=choice, value='single', bg='gray15', activebackground='gray15', fg='gray', command=button_check)
singleRadioButton.grid(row=0, column=0, padx=20)
#  Multiple Radio button 2 multiple
multipleRadioButton=Radiobutton(chooseFrame, text='Multiple', font=('times new roman', 25, 'bold'), variable=choice, value='multiple', bg='gray15', activebackground='gray15', fg='gray', command=button_check)
multipleRadioButton.grid(row=0, column=1, padx=20)
# to select the single or multiple radio button 
choice.set('single')


#To Field
toLabelFrame=LabelFrame(root, text='To (Email Address)', font=('times new roman', 16, 'bold'), bd=3, fg='white', bg='gray15')
toLabelFrame.grid(row=2, column=0, padx=140)
# Entry Fields to enter some text
toEntryField=Entry(toLabelFrame, font=('times new roman', 18, 'bold'), width=30)
toEntryField.grid(row=0, column=0)
#browse button with image
browseImage=PhotoImage(file='browse.png')
browseButton =Button(toLabelFrame, text=' Browse', image=browseImage, compound=LEFT, font=('arial', 12, 'bold'), cursor='hand2',bd=0,bg='gray15', activebackground='gray15', fg='gray', state=DISABLED, command=browse)
browseButton.grid(row=0, column=1, padx=20)

#Subject Field
subjectLabelFrame=LabelFrame(root, text='Subject', font=('times new roman', 16, 'bold'), bd=3, fg='white', bg='gray15')
subjectLabelFrame.grid(row=3, column=0, pady=10)

subjectEntryField=Entry(subjectLabelFrame, font=('times new roman', 18, 'bold'), width=30)
subjectEntryField.grid(row=0, column=0)

#Compose Email Label Frame Text Area with speak and attachment buttons
emailLabelFrame=LabelFrame(root, text='Compose Email', font=('times new roman', 16, 'bold'), bd=3, fg='white', bg='gray15')
emailLabelFrame.grid(row=4, column=0, padx=20)
#mic button
micImage=PhotoImage(file='mic.png')
Button(emailLabelFrame, text=' Speak', image=micImage, compound=LEFT, font=('arial', 12, 'bold'), cursor='hand2',bd=0,bg='gray15', activebackground='gray15', fg='gray', command=speak).grid(row=0, column=0)
# Attacment Button with png image 
attachImage=PhotoImage(file='attach.png')
Button(emailLabelFrame, text=' Attachment', image=attachImage, compound=LEFT, font=('arial', 12, 'bold'), cursor='hand2',bd=0,bg='gray15', activebackground='gray15', fg='gray', command=attachment).grid(row=0, column=1)
#Body of the Email with text area
textarea=Text(emailLabelFrame, font=('times new roman', 14 ), height=8)
textarea.grid(row=1, column=0, columnspan=2)

# 3 buttons send, clear and exit button 
# 1st button Send button
sendImage=PhotoImage(file='send.png')
Button(root,image=sendImage, bd=0, bg='gray15', cursor='hand2', activebackground='gray15', command=send_email).place(x=490, y=540)
# 2nd button Clear button
clearImage=PhotoImage(file='clear.png')
Button(root,image=clearImage, bd=0, bg='gray15', cursor='hand2', activebackground='gray15', command=clear).place(x=590, y=550)
#3rd button Exit button
exitImage=PhotoImage(file='exit.png')
Button(root,image=exitImage, bd=0, bg='gray15', cursor='hand2', activebackground='gray15', command=iexit).place(x=690, y=550)



#Labels  total, sent  Left and failed 
# total label 1
totalLabel=Label(root, font=('times new roman', 18, 'bold'), bg='gray15', fg='gray')
totalLabel.place(x=10, y=560)
# send label 2
sentLabel=Label(root, font=('times new roman', 18, 'bold'), bg='gray15', fg='gray')
sentLabel.place(x=100, y=560)
# Left label (how many email are left to sent)
leftLabel=Label(root, font=('times new roman', 18, 'bold'), bg='gray15', fg='gray')
leftLabel.place(x=190, y=560)
# filed label  (how many emails are not sent and they are failed)
failedLabel=Label(root, font=('times new roman', 18, 'bold'), bg='gray15', fg='gray')
failedLabel.place(x=280, y=560)








root.mainloop()