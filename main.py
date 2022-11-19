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
    print("hi")
    doc = docx.Document(filename)
    fulltext = []
    allvars = []
    amountofvars = 0
    currentVar = ""
    varChars = '{{'

    #Go through each paragraph in the word file
    for paragraph in doc.paragraphs:
        # Here we check each words in the file for {{}} which should contain a variable
        currentPar = paragraph.text
        index = 0
        for x in len(currentPar):
            if currentPar[index] != '{':
                index += 1
                continue
            
            index += 1
                

        
            

#We use this to edit the document
def edit_document(file_path, final_document):
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
    #context passed over to the word document
    context = {'event_name_informal': event_name_informal, 'date': date, 'target_name': target_name, 'event_name':event_name,
    'rsvp_date': rsvp_date, 'my_number': my_number, 'my_email': my_email, 'my_name': my_name}
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
                sg.popup_error("not yet impletmented")
    window.close()

#-------------CODE RUN---------------

gui_scan()
