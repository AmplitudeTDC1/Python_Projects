import win32com.client
import os

word = win32com.client.Dispatch("word.Application")
word.visible = 0

doc_pdf = "AUWAL ABUBAKAR.pdf"
input_file = os.path.join(os.path.dirname(__file__), doc_pdf)

wb = word.Documents.Open(input_file)
output_file_name = doc_pdf[:-4] + ".docx"
output_file_path = os.path.join(os.path.dirname(__file__), output_file_name)
wb.SaveAs2(output_file_path, FileFormat=16)
print(".pdf to .docx conversion is complete")
wb.Close()

word.Quit()
