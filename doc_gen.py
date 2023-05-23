from docxtpl import DocxTemplate



#Import the invoice template (has specific sintax that
# allows easy creation of invoice)
doc = DocxTemplate("invoice_template.docx")

#Generates the new invoice
doc.render({})
doc.save("new_invoice.docx")