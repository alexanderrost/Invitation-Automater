# Project for automaticly generating word files for things like invitations and coverletters based on a given template. Hopefully with UI interface.

#We use pysimplegui for the GUI element of the program
import PySimpleGUI as sg
from pathlib import Path

#We use docxtpl to edit the word template
from docxtpl import DocxTemplate
from datetime import datetime

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
    print(allvars)
    return allvars       
                

        
            

#We use this to edit the document
def edit_document(file_path, final_document):
    vars = []
    vars.append(scan_document(file_path))
    #Here goes the path to your template
    doc = DocxTemplate(file_path)
    #this will later be replaced with userinput
    event_name_informal = "Big party"
    date = datetime.today().strftime("%d/%B/%Y")
    target_name = "John"
    event_name = "Big Party at my house!"
    rsvp_date = "11/12/23"
    my_number = "(123) 456 789"
    my_email = "partyguy@gmail.com"
    my_name = "Alexander"
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
                vars = []
                vars.append(scan_document(values["-IN-"]))
                window.close()
                gui_check(vars)

    window.close()

def gui_check(vars):
    layout = []
    options = []
    options.append(vars)
    currentCheckBox = ""
    for x in range(len(vars)):
            currentCheckBox = str(options[x])
            print("Current box is: " + currentCheckBox)
            layout.append([sg.Checkbox(str(currentCheckBox), default=False)])
            currentCheckBox = ""
    
    layout.append([sg.Exit()])

    window = sg.Window("Select thingy", layout)
    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, "Exit"):
            break

    window.close()

#-------------CODE RUN---------------

# scan_document('generator_template_py.docx')
gui_scan()
