from easygui import *
from pyPdf import PdfFileReader, PdfFileWriter

# prompt user for input file
inputFileName = fileopenbox("", "Select a file to extract from...", "*.pdf")
if inputFileName == None:
    exit()

inputFile = PdfFileReader(file(inputFileName, "rb"))
numPages = inputFile.getNumPages()

# prompt user for page from which to begin selection
beginSelection = int(enterbox("Enter a beginning page for the selection: ", "Enter a beginning page..."))
if beginSelection == None:
    exit()
while beginSelection > numPages:
    msgbox("Page out of range!", "Please enter a page number less than {}".format(numPages))
    beginSelection = int(enterbox("Enter a beginning page for the selection: ", "Enter a beginning page..."))

# prompt user for page at which to end selection
endSelection = int(enterbox("Enter an ending page for the selection: ", "Enter an ending page..."))
if endSelection == None:
    exit()
while endSelection > numPages or endSelection < beginSelection:
    msgbox("Page out of range!", "Please enter a page number between {} and {}".format(beginSelection, numPages))
    endSelection = int(enterbox("Enter an ending page for the selection: ", "Enter an ending page..."))

# prompt user for output destination
outputFileName = filesavebox("", "Save the extracted pages to...", "*.pdf")
while inputFileName == outputFileName:
    msgbox("Cannot overwrite original file!", "Please select another output file...")
    outputFileName = filesavebox("", "Save the extracted pages to...", "*.pdf")
if outputFileName == None:
    exit()

# read in selection and write to output file
outputPDF = PdfFileWriter()
for pageNum in range(beginSelection-1, endSelection):
    page = inputFile.getPage(pageNum)
    outputPDF.addPage(page)
outputFile = file(outputFileName, "wb")
outputPDF.write(outputFile)
outputFile.close()
