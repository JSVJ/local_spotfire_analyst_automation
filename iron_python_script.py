# REFERENCE - https://community.tibco.com/s/article/How-to-export-an-analysis-to-PDF-using-IronPython-in-TIBCO-Spotfire
# Copyright © 2017. TIBCO Software Inc.  Licensed under TIBCO BSD-style license.

# Import namespaces
from Spotfire.Dxp.Framework.ApplicationModel import ApplicationThread
from Spotfire.Dxp.Application.Export import PdfExportSettings
from Spotfire.Dxp.Application.Export import ExportScope
from Spotfire.Dxp.Application.Export import PageOrientation
from Spotfire.Dxp.Application.Export import PaperSize
from Spotfire.Dxp.Application.Export import ExportScope
import time
import clr
import datetime
from System import DateTime

# RELOAD ALL THE DATA (IF NEEDED)
Document.Data.Tables.ReloadAllData()

# Declaring the function which will run async
def g(app,fileName,pdfexpsettings):
   def f():      
	  app.Document.Export(pdfexpsettings,fileName)
   return f

# Set the file name and page to export
PAGE = Document.Pages[0]
FILENAME = 'TEST'
FILELOCATION = 'D:\\'

today = DateTime.Now
DATE = str(today.Year) + '_' + str(today.Month) + '_' + str(today.Day) + '__' + str(today.Hour) + '_' + str(today.Minute) + '_' + str(today.Second)
fileName = FILELOCATION + FILENAME + '_' + DATE + ".pdf"


Document.ActivePageReference = PAGE

# LOOPING THROUGH ALL THE PAGES
# for page in Document.Pages:
	# Document.ActivePageReference = page

pdfexpsettings = PdfExportSettings()
pdfexpsettings.Scope = ExportScope.AllPages
pdfexpsettings.PageOrientation = PageOrientation.Landscape
pdfexpsettings.IncludePageTitles = True
pdfexpsettings.IncludeVisualizationTitles = True
pdfexpsettings.IncludeVisualizationDescriptions = False
pdfexpsettings.PaperSize = PaperSize.A4
pdfexpsettings.PageOrientation = PageOrientation.Landscape

# Executing the function on the application thread, and Save the document back to the Library
Application.GetService[ApplicationThread]().InvokeAsynchronously(g(Application, fileName,pdfexpsettings))

# Note:
# The function g is necessary because the script's scope is cleared after execution,
# and therefore Application (or anything else defined in this scope) will not be available
# when the code invokes on the application thread.



# DO NOT UNCOMMENT THE BELOW LINES. 
# THEY ARE AFFECTING THE REPORT
# time.sleep(3)


# clr.AddReference("System.Windows.Forms")
 
# from System.Windows.Forms import MessageBox
# MessageBox.Show("Good Day Sir,\n\n\nPDF FILE HAS BEEN SAVED.\nNOW TRYING TO CLOSE THE APPLICATION\n\n See you.")