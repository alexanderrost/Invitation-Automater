# Project for automaticly generating word files for things like invitations and coverletters based on a given template. Hopefully with UI interface.
# By Alexanderrost 2022


#TODO Fix the excel scanning and data insertion functions once checkboxes work reliabily, then we can finalize the project --> Highest prio
#TODO Fix layouts and the design of the program, looks like a toddler drew it. Also make it an .exe?
#TODO Write tests and make sure they work before first github release.
#TODO CLEAN UP THE CODE AND ADD MORE RELEVANT COMMENTS, also use more standardized language in the code(look up PEP 8 as recommended by a friend)
#TODO Fix README.md file and clean up the github a little once it's all donezo.
#TODO Clean up the scanning document function.

#We use pysimplegui for the GUI element of the program
#We might change this out for customtkinter once I get it working.
import PySimpleGUI as sg
from pathlib import Path

#We use docxtpl to edit the word template
from docxtpl import DocxTemplate

#Docx will be used to scan our file for the variables enterd by the user
import docx



#Used to combat errors when users didn't specifiy a filepath
def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    sg.popup_error("Filepath not correct")
    return False

#Here we will find what the user wants us to fill in using our code
def scan_document(filename):
    # This will scan the document for any part of the text that has "{{name_of_input_here}}"
    # Like the template file has {{event_name_informal}} 
    doc = docx.Document(filename)
    allvars = []
    amountofvars = 0
    currentVar = ""
    #Checks if we are at the first of the two {{ indicating there is a variable to come
    firstSquiggly = True
    countChars = False

    #Go through each paragraph in the word file
    for paragraph in doc.paragraphs:
        # Here we check each words in the file for {{}} which should contain a variable
        currentPar = paragraph.text
        for x in range(len(currentPar)):
            if countChars == True and currentPar[x] != '}':
                currentVar = currentVar + currentPar[x]
                continue
            else:
                countChars = False
                if currentPar[x] == '{' and firstSquiggly == True:
                    allvars.append(currentVar)
                    amountofvars += 1
                    currentVar = ""
                    firstSquiggly = False
                    continue
                elif currentPar[x] == '{' and firstSquiggly == False:
                    countChars = True
                    firstSquiggly = True

    #Needs fixing later on. First element is wrong but im lost
    amountofvars -= 1
    allvars.remove("")
    return allvars       
                

        
            

#We use this to edit the document
def edit_document(file_path, final_document):
    vars = []
    vars.append(scan_document(file_path))
    #Here goes the path to your template
    doc = DocxTemplate(file_path)
    context = {}
    #context passed over to the word document
    for x in range(len(vars)):
        context[vars[x]] = ""
    # Render and save the document at specified filepath
    doc.render(context)
    print("Document completed")
    doc.save(final_document)

#Function used for most/all of our gui needs
def gui_scan():
    sg.theme("Python")
    #The layout is for what elements our GUI has
    #With the option to browse for you template file aswell as pick a final destination for your completed file
    layout = [
        [sg.Text("Template file:"), sg.Input(key="-IN-"), sg.FileBrowse(file_types=(("Word files", "*.docx"),))],
        [sg.Exit(), sg.Button("Scan file")],
    ]

    #Creates the window based on our layout
    window = sg.Window("Invitation Generator", layout)
    #This loops check what's happening with the gui at all times, I know while True is awful
    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, "Exit"):
            break
        if event == "Scan file":
            #check if the user has enterd a valid filepath
            if is_valid_path(values["-IN-"]):
                #This is a really ugly solution, I really have to fix this later on :) but ive spent an ungodly amount of time getting this to work
                vars = []
                tempvars = []
                tempvars = scan_document(values["-IN-"])
                for x in range(len(tempvars)):
                    vars.append(tempvars[x])
                window.close()
                gui_check(vars)

    window.close()

def gui_check(vars):
    layout = []
    layout.append([sg.Text('Choose the options you DO NOT want to be dynamic', key='-TEXT-')])
    currentCheckBox = ""
    for x in range(len(vars)): 
            currentCheckBox = str(vars[x])
            layout.append([sg.CB("" + currentCheckBox, default=False, key= "-CB" + str(x) + "-")])
            currentCheckBox = ""
    
    layout.append([sg.Exit()])

    window = sg.Window("Dynamic options", layout, resizable=True)
    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, "Exit"):
            break

    window.close()

#This file will be used to scan excel files to determin what data goes to what variable
def gui_xcel_check():
    return 0

# This will take the data from the forms and put it in the right place on the invitations
def gui_insert_data():
    return 0



#-------------CODE RUN---------------

# scan_document('generator_template_py.docx')
gui_scan()
