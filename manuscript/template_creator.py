# This file uses Python docx to generate a template file for Pandoc based off the manuscript. 
from docx import Document
from docx.shared import Pt, RGBColor
import docx
import yaml

doc_template = Document()


# Establish styles of template document: 
doc_template.styles["Normal"].font.name = 'Calibri'
doc_template.styles["Normal"].font.size = Pt(12)

doc_template.styles["Body Text"].font.name = 'Calibri'
doc_template.styles["Body Text"].font.size = Pt(12)

headers = ["Heading 1", "Heading 2", "Heading 3"]

for h in headers: 
    doc_template.styles[h].font.name = 'Calibri'
    doc_template.styles[h].font.size = Pt(12)
    doc_template.styles[h].font.all_caps = True
    doc_template.styles[h].font.bold = False
    doc_template.styles[h].font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    
    doc_template.styles[h].next_paragraph_style = doc_template.styles["Body Text"]
    
doc_template.styles["Heading 2"].font.italic = True
doc_template.styles["Heading 3"].font.italic = True
doc_template.styles["Heading 3"].font.all_caps = False

# If details not present, create a details file


# Read in information from details file, and add the proper text. 

## Put running head into the footer. 


# Save the resulting file

doc_template.save("manuscript/demo.docx")



