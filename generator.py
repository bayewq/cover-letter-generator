from spire.doc import *
from spire.doc.common import *
from datetime import date

# Create a Document object
document = Document()
# Load a Word docx or doc document
document.LoadFromFile("CoverLetter.docx")

exactCompanyName = input("Exact Company Name: ")
simpleCompanyName = input ("Simple/Shortened Company Name")
companyAddress = input("Company Location/Address")
exactJobTitle = input("Exact Job Title: ")
simpleTitle = input("Simple Job Title: ")
foundThrough = input("How did you find this job?: ")



todayStr = str(date.today())

document.Replace("[Date]", todayStr, False, False)

# Save the resulting document
document.SaveToFile("GENERATED.docx", FileFormat.Docx2016)
document.Close()