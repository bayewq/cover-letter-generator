from spire.doc import *
from spire.doc.common import *
from datetime import date

# Create a Document object
document = Document()
# Load a Word docx or doc document
document.LoadFromFile("CoverLetter.docx")

exactCompanyName = input("\nExact Company Name: ")
simpleCompanyName = input ("Simple/Shortened Company Name: ")
if simpleCompanyName == "":
    simpleCompanyName = exactCompanyName
companyLocation = input("Company Location/Address: ")
exactJobTitle = input("Exact Job Title: ")
simpleTitle = input("Simple Job Title: ")
foundThrough = input("How did you find this job?: ")
impressedWith = input("I was impressed with... ")


todayStr = str(date.today())

document.Replace("[Date]", todayStr, False, False)
document.Replace("[Exact Company Name]", exactCompanyName, False, False)
document.Replace("[Company Location]", companyLocation, False, False)
document.Replace("[Simple Company Name]", simpleCompanyName, False, False)
document.Replace("[Exact Job Title]", exactCompanyName, False, False)
document.Replace("[Simple Title]", simpleTitle, False, False)
document.Replace("[Found Through]", foundThrough, False, False)
document.Replace("[Impressed With]", impressedWith, False, False)


# Save the resulting document
document.SaveToFile("GENERATED - Delete Warning.docx", FileFormat.Docx2016)
document.Close()