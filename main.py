# Project for automaticly generating word files for things like invitations and coverletters based on a given template. Hopefully with UI interface.

#We use docxtpl to edit the word template
from docxtpl import DocxTemplate
from datetime import datetime
#Here goes the path to your template
doc = DocxTemplate('generator_template_py.docx')

event_name_informal = "Big party"
date = datetime.today().strftime("%d/%B/%Y")
target_name = "John"
event_name = "Big Party at my house!"
rsvp_date = "11/12/23"
my_number = "(123) 456 789"
my_email = "partyguy@gmail.com"
my_name = "Alexander"

context = {'event_name_informal': event_name_informal, 'date': date, 'target_name': target_name, 'event_name':event_name,
'rsvp_date': rsvp_date, 'my_number': my_number, 'my_email': my_email, 'my_name': my_name}

doc.render(context)
print("Document printed")
doc.save("New_document_generated.docx")

