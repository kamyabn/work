#!/usr/bin/env python3

import sys
import argparse
import csv
import docx
from docx import Document
from docxtpl import DocxTemplate
import jinja2




def main():

    # Get the number of variables used in pdf from the user and pass it to function questions
    val_num = input("How many variables are contained in this document? ")
    questions(val_num)



# Puts variables and the corresponding question to the variable in a dictionary
def questions(val_num):

    number = int(val_num)
    var_arr = []
    que_arr = []
    
    for i in range(number):
        val = input("What is the next variable \n")
        #print(val)
        var_arr.append(val)  
        question = input("What does " + val + " represent \n")
        #print(question)
        que_arr.append(question)

    dictionary = dict(zip(var_arr, que_arr))

    print("Document is ready for use\n")
    state = input("Do you want to go ahead? (Y/N)\n")

    
    if (state == 'Y'):
        replace(dictionary)

    if (state == 'N'):
        print(state)
    

# This function replaces the question(value of dict) with the answer, and passes the new Dict to function final
# Ex: {'$fullname' : What's the name} ==> {'fullname': 'Kamyab Navirian'}
def replace(dictionary):

    for key in dictionary:
        answer = input(dictionary[key] + " for this specific document? ")
        dictionary[key] = answer
        #print(key)
        #print(dictionary[key])
    final(dictionary)


# This function replaces the variable in the document using new Dict, and creates another word document
def final(dictionary):

    print(dictionary)
    doc = DocxTemplate("template.docx")
    doc.render(dictionary)
    print(doc.render(dictionary))
    doc.save("output.docx")


'''
doc = docx.Document("last.docx")
all_paras = doc.paragraphs
print(len(all_paras))


for para in all_paras:
    if '$fullname' in para.text:
        print(para.text)
        para.text = 'Kamyab'
        print(para.text)


for para in all_paras:
    print(para.text)
    print("-------")
'''

if __name__ == "__main__":
   main()
