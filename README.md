# String-search-in-a-folder-either-text-or-docx-files-and-retrieve-required-data-from-each-file

#1.to get the name from excel from each row
#2.to get the file from folder which contains the name from excel
#3. then attach the email,contact number and filename to the name in excel


import os
import re
import shutil
from os import listdir
from os.path import isfile, join
import pandas as pd
from docx import Document
import docx
import xlrd

src = 'E:\Python_Udemy_Course\Resumes'
#data = pd.read_csv("Names1.csv")
input_data = input('Enter the file:')

if len(listdir(src))!=0:
   emails = []
   contact_no = []
   Document_name = []
   if input_data.find('.xlsx')!=-1:
      data = pd.read_excel(input_data)
      data.to_csv('csvfile.csv', encoding='utf-8', index=False)

      for x in range(0,len(data.values)):
          name_match = (data.values[x,0])
          for filename in [f for f in listdir(src) if isfile(join(src, f))]:
              if filename.find('.docx')!=-1:
                 #print('Hi in word')
                 x = Document(os.path.join(src, filename))
                 wholedoc=""
                 for paragraph in x.paragraphs:
                     wholedoc += paragraph.text
                 if (name_match.lower() in wholedoc.lower()):
                     Document_name.append(filename)
                     for i in range(len(x.paragraphs)):
                         email_match = re.findall(r'[\w\.-]+@[\w\.-]+',x.paragraphs[i].text)
                         if email_match:
                            emails.extend(email_match)
                         contact_match = re.findall(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]', x.paragraphs[i].text)
                         if contact_match:
                            contact_no.extend(contact_match)
              elif filename.find('.txt')!=-1:
                   x = open(os.path.join(src, filename),'r',encoding="utf8")
                   y = x.read()
                   content = y.splitlines()
                   if name_match.lower() in y.lower():
                      file_name = filename
                      Document_name.append(file_name)
                      for j in range(0,len(content)):
                          email_match = re.findall(r'[\w\.-]+@[\w\.-]+',content[j])
                          if email_match:
                             emails.extend(email_match)
                          contact_match = re.findall(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]', content[j])
                          if contact_match:
                             contact_no.extend(contact_match)

   elif input_data.find('.csv')!=-1:
        data = pd.read_csv(input_data)
        emails = []
        contact_no = []
        Document_name = []
        for x in range(0,len(data.values)):
            name_match = (data.values[x,0])
            for filename in [f for f in listdir(src) if isfile(join(src, f))]:
                if filename.find('.docx')!=-1:
                   #print('Hi in word')
                   x = Document(os.path.join(src, filename))
                   wholedoc=""
                   for paragraph in x.paragraphs:
                       wholedoc += paragraph.text
                   if (name_match.lower() in wholedoc.lower()):
                      Document_name.append(filename)
                      for i in range(len(x.paragraphs)):
                          email_match = re.findall(r'[\w\.-]+@[\w\.-]+',x.paragraphs[i].text)
                          if email_match:
                             emails.extend(email_match)
                          contact_match = re.findall(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]', x.paragraphs[i].text)
                          if contact_match:
                             contact_no.extend(contact_match)
                elif filename.find('.txt')!=-1:
                     x = open(os.path.join(src, filename),'r',encoding="utf8")
                     y = x.read()
                     content = y.splitlines()
                     if name_match.lower() in y.lower():
                        file_name = filename
                        Document_name.append(file_name)
                        for j in range(0,len(content)):
                            email_match = re.findall(r'[\w\.-]+@[\w\.-]+',content[j])
                            if email_match:
                               emails.extend(email_match)
                            contact_match = re.findall(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]', content[j])
                            if contact_match:
                               contact_no.extend(contact_match)

   else:
        print('No input file is entered')
   data['emails'] = emails
   data['Contact_No'] = contact_no
   data['Document_name'] = Document_name
   data.to_csv('Append_Docx.csv')             
else:
      print('No files found in directory')         
        
