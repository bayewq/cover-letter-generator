from spire.doc import *
from spire.doc.common import *
import time

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
commitedTo = input("What is this company committed to?")
culture = input("what culture does this company exhibit?")
roleDetails = input("What about this role excites you?")
whyRole = input ("Why does that excite you?")

todayStr = time.strftime("%B %d, %Y")

document.Replace("[Date]", todayStr, False, False)
document.Replace("[Exact Company Name]", exactCompanyName, False, False)
document.Replace("[Company Location]", companyLocation, False, False)
document.Replace("[Simple Company Name]", simpleCompanyName, False, False)
document.Replace("[Exact Job Title]", exactJobTitle, False, False)
document.Replace("[Simple Title]", simpleTitle, False, False)
document.Replace("[Found Through]", foundThrough, False, False)
document.Replace("[Committed To]", commitedTo, False, False)
document.Replace("[Culture]", culture, False, False)
document.Replace("[Role Details]", roleDetails, False, False)
document.Replace("[Why This Role]", whyRole, False, False)


# Save the resulting document
document.SaveToFile("GENERATED - Delete Warning.docx", FileFormat.Docx2016)
document.Close()
