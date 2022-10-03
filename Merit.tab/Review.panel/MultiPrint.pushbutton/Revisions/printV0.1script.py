#from asyncio.windows_events import NULL
import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *
# Import ToDSType(bool) extension method
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
# Import geometry conversion extension methods
clr.ImportExtensions(Revit.GeometryConversion)
# Import DocumentManager and TransactionManager
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *
# Import RevitAPI
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

import math 
import traceback

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import Application as dapp
from System.Windows.Forms import Form, TextBox, Label, Button, MessageBox, CheckBox
from System.Windows.Forms import ToolBar, ToolBarButton, OpenFileDialog, FolderBrowserDialog, ToolStripButton, ToolStrip
from System.Windows.Forms import DialogResult, ScrollBars, DockStyle, ListView, ListViewItem, ColumnHeaderAutoResizeStyle
from System.Windows.Forms.Form import Size
from System.Drawing import Point, Image

from pyrevit import revit, DB, UI

from operator import itemgetter

doc = __revit__.ActiveUIDocument.Document
#docTitle = doc.Title
uidoc = __revit__.ActiveUIDocument
#uiapp = DocumentManager.Instance.CurrentUIApplication
#app = uiapp.Application

doc1=__revit__.ActiveUIDocument

app = __revit__.Application








itemlist = ["item1","item2"]


#_____________________Sheet filtering and selection ________________

# gets all the sheets from the current document 
sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).\
                    WhereElementIsNotElementType().ToElements()

#gets all the printer settings saved 
printSettingCollector = FilteredElementCollector(doc).OfClass(PrintSetting)


SHEETList=[]
for sheet in sheets:
	sheetId = sheet.Id
	sheetNumber = sheet.SheetNumber
	sheetName = sheet.Name
	sheetFN = sheetNumber + '_' + sheetName
	SHEETList.append(sheetFN)
	#print(sheetFN)
	#print(dir(sheet))
	#print(sheetName)
	#print(sheetId)
	#print('{}  {}. '.format(sheetNumber, sheetName))
	
PAM_no_OG = 'Enter PAM Number Original'
run = False
checkedPdf = False
i = 0
viewItemList = []
checkedlist =[]
sheetlist=[]
selectedsheetSize = ""
pdf_directory = ""
errorReport=[]

class MyClass(Form):

	def __init__(self):
		self.Size = Size.PropertyType(800,1000) #sets the size of the gui 
		global PAM_no_OG
		global itemlist
		global sheetFN
		global i
		global viewItemList
		#global Frame_no
		name = 0
		self.Text = 'User Input'
		self.label = Label()
		self.label.Text = ""
		self.label.Location = Point(650, 200)
		#self.label.BackColor = Label.BackColor.PropertyType.Red
		#++____________ Image add
		img = Image.FromFile("C:\Users\youssef.elshebani\Downloads\logo-merit.png")
		self.label.Image = img
		
		
		
		self.label.Height = img.Height
		self.label.Width = img.Width
		self.label.AutoSize = False

		#________________________________
		self.label2 = Label()
		self.label2.Text = 'Browse to the directory and tick\n the sheets required for printing'
		self.label2.Location = Point(50,50)
		self.label2.AutoSize = True
		self.Controls.Add(self.label2)

		#______________Text Box _________________
		self.textbox = TextBox()
		self.textbox.Text = str(PAM_no_OG)
		self.textbox.Enabled = True
		self.textbox.Location = Point(700, 50)
		self.textbox.Width = 200
		self.textbox.Height = 100
		self.Controls.Add(self.label)
		#self.Controls.Add(self.textbox)
		#self.Controls.Add(self.textbox.Text)
		#self.all_c.append(self.textbox.Text)
		#PAM = self.textbox.Text
		self.but = Button()
		self.but.Text = 'Ok'
		self.but.Location = Point(600,700)
		self.but.Click += self.OnClick
		self.Controls.Add(self.but)
		
		#_______________ Browse Button _______________________
		self.BrowseButton = Button()
		self.BrowseButton.Text = 'Browse'
		self.BrowseButton.Location = Point(400,100)
		self.BrowseButton.Click += self.BrowseClick
		self.Controls.Add(self.BrowseButton)
		
		#________Textbox Browse tree_________________________
		
		self.TextBox2 = TextBox()
		self.TextBox2.Location = Point(100,100)
		self.TextBox2.Width = 300
		self.Controls.Add(self.TextBox2)
		#_____________ checkbox  _____________________________
		self.CheckBox1 = CheckBox()
		self.CheckBox1.Location = Point(650,100)
		self.CheckBox1.AutoSize = True
		self.CheckBox1.Width = 100
		self.CheckBox1.Text = "export PDF"
		self.Controls.Add(self.CheckBox1)

    	#_________________ ListView______________________________
		self.listView1 = ListView()
		self.listView1.Location = Point(50,150)
		self.listView1.Scrollable = True
		self.listView1.Height = 600
		self.listView1.Width = 500
		self.listView1.Columns.Add("Item")
		self.listView1.View = ListView.View.PropertyType.Details		
		self.listView1.CheckBoxes = True
		for count, sheet in enumerate(SHEETList):
			#print(SHEETList[i])
			s=sheet
			s = ListViewItem(str(s),0) # viewitem 
			s.Text = str(SHEETList[count])
			
			viewItemList.append(s)
			#self.listviewItem1.Text = "Here"
			#self.listviewItem1.Checked = False
			#collection = self.listView1.ListViewItemCollection
			#collection.Add(self.listviewItem1,"item")
			#collect = self.listView1.ListViewItemCollection.Add(self.listviewItem1,"")
		
		
			self.listView1.Items.Add(s)
		self.listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
		self.listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)

    	
		self.Controls.Add(self.listView1)
    
    
	def BrowseClick(self,sender,args):
		dialog = FolderBrowserDialog()
		
		if dialog.ShowDialog(self) == DialogResult.OK:
			global pdf_directory
			pdf_directory = dialog.SelectedPath
			self.TextBox2.Text = str(pdf_directory)

    
	def OnClick(self,sender,args):
		MessageBox.Show(self.textbox.Text)
		global PAM_no_OG
		global run
		global viewItemList	
		global sheetlist
		global checkedPdf
		global checkedlist
		PAM_no_OG = self.textbox.Text
		if self.CheckBox1.Checked:
			checkedPdf = True
		for count,item in enumerate(viewItemList):
			if item.Checked:
				checkedlist.append(item)
				sheetlist.append(sheets[count])
				
		
		#run = self.CheckBox1.Checked
		self.Close()
		#def Close(self):
		#self.close()
		#form = LabelDemoForm()

def sheetSize(sheet):
	global selectedsheetSize
	try:
		titleblock = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).\
												WhereElementIsNotElementType().ToElements()
		sheetpara = [i for i in titleblock]
		
		for parameter in sheetpara:
		
			sheetHeight = parameter.LookupParameter("Sheet Height").AsValueString()
			sheetWidth = parameter.LookupParameter("Sheet Width").AsValueString()
			if sheetHeight ==  "594":
				selectedsheetSize = "A1"
				#break
			elif sheetHeight == "297":
				selectedsheetSize = "A3"
				#break
			elif sheetHeight == "210":
				selectedsheetSize = "A4"
				#break
			else:
				selectedsheetSize = "A0"
		#global sheetSize
		return selectedsheetSize, sheetHeight,sheetWidth
	except:
		import traceback
		return traceback.format_exc()
	


def printSetSelect(name): # A function that looks for a specific print name and returns 
	# a printsetting type that matches that name 
	for currentprintselect in printSettingCollector:
		if currentprintselect.Name == name:
			return currentprintselect











dapp.Run(MyClass())
#print(sheetlist)
#print(pdf_directory)
"""
sizesheet = sheetSize(sheetlist[0])
print(sizesheet)




"""






printManager = doc.PrintManager #insitiating the print manager from the current document 

for selectsheet in sheetlist: # iterate through the selected sheets for printing 
	selectedSheetSize = str(sheetSize(selectsheet)[0]) #extracts the sheet size from the selected sheet
	#print(selectedSheetSize)
	MessageBox.Show(str(sheetlist))
	if checkedPdf:
		transaction1 = Transaction(doc)
		
		




		


		try:
			transaction1.Start('print')


			viewset = ViewSet()  #init a new instance of the ViewSet class
			viewset.Insert(selectsheet)


			revision = Revision.Create(doc)
			revision.Description = "For Faprication"
			revision.IssuedTo = ''
			revision.NumberType = RevisionNumberType.Numeric
			r_id = revision.Id
			
			printManager.PrintRange = PrintRange.Select
			printManager.Apply()
			printManager.SelectNewPrintDriver("PDFCreator") #select the print driver 
			printManager.Apply()
			printManager.CombinedFile = False 
			printManager.Apply()
			
			if printManager.IsVirtual: #check if the driver selected is a virtual driver 
				printManager.PrintToFile = True #sets the print to file method to true to enable 
				printManager.Apply()
			else:
				printManager.PrintToFile = True
				printManager.Apply()
				
		#Setting the file path by extracting parameters from the sheet and used to name the pdf file 
		#that will be created 
		
			pn = selectsheet.LookupParameter("MRT_SHT_Project Number").AsString()
			m_rt = selectsheet.LookupParameter("MRT_SHT_MRT").AsString()
			zn = selectsheet.LookupParameter("MRT_SHT_Zone").AsString()
			levl = selectsheet.LookupParameter("MRT_SHT_Level").AsString()
			doct = selectsheet.LookupParameter("MRT_SHT_Document Type").AsString()
			rle = selectsheet.LookupParameter("MRT_SHT_Role").AsString()
			sn = selectsheet.LookupParameter("MRT_SHT_Sheet").AsString()
			cr =  selectsheet.LookupParameter("Current Revision").AsString()
			t1 = selectsheet.LookupParameter("MRT_SHT_Title 1").AsString()
			t2 = selectsheet.LookupParameter("MRT_SHT_Title 2").AsString()
			t3 = selectsheet.LookupParameter("MRT_SHT_Title 3").AsString()


			FilePathParameterlist = [pn, m_rt, zn, levl, doct, rle, sn, cr, t1, t2, t3]

			#NoneTypeFinder = [iNone for iNone,val in enumerate(FilePathParameterlist) if val == None]
			
			if None in FilePathParameterlist:
				MessageBox.Show("at least one sheet parameter is missing, recommedn check and reprint", "Error Message")

			#fileterd parameter list if any of the parameter is none its discarded from the name
				FPL = [ele for ele in FilePathParameterlist if ele != None]
				outpath = pdf_directory + "\\"
				for ele in FPL:
					outpath = outpath + ele + "-"
				outpath = outpath + ".pdf"
			else:
				outpath = pdf_directory + "\\" + pn + "-" + m_rt + "-" + zn + "-" + levl + "-" + doct + "-" + rle + "-" + sn + "_" + cr + '_' + t1 + " " + t2 + " " + t3 + ".pdf"	

			
			
			
			printManager.PrintToFileName = outpath #set the file name when printing 
			printManager.Apply()
			
			
			
			viewSheetSetting = printManager.ViewSheetSetting 
			viewSheetSetting.CurrentViewSheetSet.Views = viewset
			printManager.Apply()
			
			
			#_________________ Print setup_________________

			
			
			printSetup = printManager.PrintSetup
			#printsetting = PrintSetting()
			
			papersizes = printManager.PaperSizes

			for papersize in papersizes:
				if papersize.Name.ToString() == selectedSheetSize:
			
					printSetup.CurrentPrintSetting.PrintParameters.PaperSize = papersize
			printSetup.CurrentPrintSetting.PrintParameters.ZoomType = ZoomType.Zoom
			printSetup.CurrentPrintSetting.PrintParameters.Zoom = 100
			if sheetSize(selectsheet)[2] > sheetSize(selectsheet)[1]:
				printSetup.CurrentPrintSetting.PrintParameters.PageOrientation = PageOrientationType.Landscape
			else:
				printSetup.CurrentPrintSetting.PrintParameters.PageOrientation = PageOrientationType.Portrait 
			printSetup.CurrentPrintSetting.PrintParameters.PaperPlacement = PaperPlacementType.Center
			printSetup.CurrentPrintSetting.PrintParameters.HiddenLineViews = HiddenLineViewsType.VectorProcessing
			printManager.Apply()
			
			printSettingName = "Print_Profile_1"
			printSetup.SaveAs(printSettingName)
			
			#
			printSetup.CurrentPrintSetting = printSetSelect(printSettingName)
			
				#errorReport.append(traceback.format_exc())
				
				

			doc.Regenerate
			printManager.SubmitPrint()
			doc.Regenerate
			
			printSetup.Delete()
			doc.Regenerate
			transaction1.Commit()
		except Exception as exp:
			if (transaction1 != None):
				transaction1.RollBack()
			MessageBox.Show(exp.message,'Error Message')
		#errorReport.append(traceback.format_exc())
		#print("where am I")
		
	else:
		MessageBox.Show("PDF Export is disabled")
		continue	

		
		




	save_options = SaveAsOptions()
	save_options.OverwriteExistingFile = True
	save_options.MaximumBackups = 1
	
	
