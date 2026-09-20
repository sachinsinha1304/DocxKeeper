from docx import Document

doc = Document()
doc.add_paragraph("hi")
doc.save("./documents/CoreTech/team.docx")