"Tag welds"
#from pickle import FALSE
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

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import Application as dapp
from System.Windows.Forms import Form, TextBox, Label, Button, MessageBox, CheckBox
from System.Windows.Forms import ToolBar, ToolBarButton, OpenFileDialog, FolderBrowserDialog, ToolStripButton, ToolStrip
from System.Windows.Forms import DialogResult, ScrollBars, DockStyle

from System.Drawing import Point

from pyrevit import revit, DB, UI

from operator import itemgetter

doc = __revit__.ActiveUIDocument.Document
#docTitle = doc.Title
uidoc = __revit__.ActiveUIDocument
#uiapp = DocumentManager.Instance.CurrentUIApplication
#app = uiapp.Application

doc1=__revit__.ActiveUIDocument

app = __revit__.Application




PAM_no_OG = 'Enter PAM Number Original'
run = False









class MyClass(Form):

	def __init__(self):
		global PAM_no_OG
		#global Frame_no
		name = 0
		self.Text = 'User Input'
		self.label = Label()
		self.label.Text = ""
		self.label.Location = Point(100, 150)
		self.label.Height = 700
		self.label.Width = 700
		self.textbox = TextBox()
		self.textbox.Text = str(PAM_no_OG)
		self.textbox.Enabled = True
		self.textbox.Location = Point(50, 50)
		self.textbox.Width = 500
		self.textbox.Height = 700
		self.Controls.Add(self.label)
		self.Controls.Add(self.textbox)
		#self.Controls.Add(self.textbox.Text)
		#self.all_c.append(self.textbox.Text)
		#PAM = self.textbox.Text
		self.but = Button()
		self.but.Text = 'Ok'
		self.but.Location = Point(150,125)
		self.but.Click += self.OnClick
		self.Controls.Add(self.but)
		#_____________ Textbox 2 _____________________________
		self.CheckBox1 = CheckBox()
		self.CheckBox1.Location = Point(150,100)
		self.CheckBox1.AutoSize = True
		self.CheckBox1.Width = 100
		self.CheckBox1.Text = "Run"
		self.Controls.Add(self.CheckBox1)

    
    
    
    
    
    
	def OnClick(self,sender,args):
		MessageBox.Show(self.textbox.Text)
		global PAM_no_OG
		global run
			
		PAM_no_OG = self.textbox.Text
		run = self.CheckBox1.Checked
		self.Close()
		#def Close(self):
		#self.close()
		#form = LabelDemoForm()
dapp.Run(MyClass())





# definign the the transaction instance 
trans = Transaction(doc)  


#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PART 1 - DEFINITIONS ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# --- MISC DEFINITIONS ---

utid = UnitTypeId.Millimeters

def conv(x):
	c = UnitUtils.ConvertToInternalUnits(x,utid)
	return c
	
def change_name(x,n):
	param_viewname = x.get_Parameter(BuiltInParameter.VIEW_NAME)
	clone_name = x.Name
	new_name_1 = clone_name.replace(str(PAM_no_OG),str(n))
	new_name_2 = new_name_1.replace("Copy 1", " ")
	param_viewname.Set(new_name_2)

def prefix(assembly_id):
	el = doc.GetElement(assembly_id)
	n = el.Name
	return n

def sheet_name(s,n):
	param_name = s.get_Parameter(BuiltInParameter.SHEET_NAME)
	param_number = s.get_Parameter(BuiltInParameter.SHEET_NUMBER)
	number_OG = sheet.SheetNumber
	new_number = str(number_OG).replace(str(PAM_no_OG),str(n))
	param_number.Set(new_number)
	new_name = str(sheet.Name)
	param_name.Set(new_name)
	return new_number


#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PART 2 - GET SHEET & VIEWS ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# --- GET SHEET ---

sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).\
													WhereElementIsNotElementType()

sheet_no_list = []
PAM_spool_sheets = []

for s in sheets:
	sheet_no = s.SheetNumber
	if sheet_no.Contains(str(PAM_no_OG) + "_"):
	#if sheet_no == str(PAM_no_OG):
		#unwrapped = UnwrapElement(s)
		PAM_spool_sheets.append(s)

"""	
PAM_spool_sheets = []

for p in PAM_sheets:
	s_name = p.Name
	if s_name.Contains("Spool"):
		PAM_spool_sheets.append(doc.GetElement(p))
		
"""		
TESTER = PAM_spool_sheets[0].GetAllViewports()
# --- GET SPOOL VIEWS ---

views_OG = []
spool_views = []
for p in PAM_spool_sheets:
	viewport_ids = p.GetAllViewports()
	viewport_names = []
	
	
	for i in viewport_ids:
		v = doc.GetElement(i)
		v_id = v.ViewId
		v_el = doc.GetElement(v_id)
		v_name = v_el.Name
		views_OG.append(v)
		viewport_names.append(v_name)
		spool_views.append(v_el)
		#if v_name.Contains(" - CA"):
			#spool_01.append(v_el)
		"""	
		matches = [" - CA", " - N2", " - CO2", " - O2"]
		if any(x in v_name for x in matches):
			spool_views.append(v_el)"""
	
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PART 3 - ELEMENTS & CATEGORIES ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# --- PIPE ACCESSORY/ FITTING CATEGORIES ---


categories = [i for i in doc.Settings.Categories]
category_name = []
category_search = []

cat_pipe_a_tags = ()
cat_pipe_f_tags = ()

cat_pipe_a = ()
cat_pipe_f = ()

for c in categories:
	cat_name = c.Name
	category_name.append(cat_name)
	if cat_name == "Pipe Accessory Tags":
		cat_pipe_a_tags = c
	if cat_name == "Pipe Fitting Tags":
		cat_pipe_f_tags = c
	if cat_name == "Pipe Accessories":
		cat_pipe_a = c
	if cat_name == "Pipe Fittings":
		cat_pipe_f = c

# --- GET PIPE FITTINGS ---


tagz = []
test = []


for n in range(len(spool_views)):
	numero = 1
		
	p_fittings = FilteredElementCollector(doc, spool_views[n].Id).WhereElementIsNotElementType().ToElements()
	p_types = []
	p_ids = []
	p_names = []
	p_locate = []
	p_filtered = []
	test.append(p_fittings)	
	for p in p_fittings:
		#UnwrapElement(p)
		p_type = p.GetType()
		p_id = p.Id	
		
		if str(p_type).Contains("FamilyInstance"):
			if p.Category.Id == cat_pipe_a.Id:
				p_types.append(p_type)
				p_ids.append(p_id)
				p_filtered.append(p)
				p_l = p.Location.Point
				p_locate.append(p_l)
			elif p.Category.Id == cat_pipe_f.Id:
				p_types.append(p_type)
				p_ids.append(p_id)
				p_filtered.append(p)
				p_l = p.Location.Point
				p_locate.append(p_l)
			
	point_map = map(lambda p: [p.X + p.Y + p.Z, p], p_locate)
	
	map_0 = map(lambda x: x[0], point_map)
	map_1 = map(lambda x: x[1], point_map)
		
	
	zip_01 = zip(map_0,map_1,p_filtered)
	sortz = sorted(zip_01, key = itemgetter(0))
	#sortz = sorted(zip_01, key = itemgetter(0 + 1))

	sorted_elements = map(lambda x: x[2],sortz)
	sorted_locations = map(lambda x: x[1], sortz)
	
	# --- GET WELD NUMBER PARAMETERS ---

	sorted_pipe_fittings = [i for i in sorted_elements if i.Category.Id == cat_pipe_f.Id]
	sorted_pipe_accessories = [i for i in sorted_elements if i.Category.Id == cat_pipe_a.Id]
	sorted_elements2 = sorted_pipe_fittings + sorted_pipe_accessories
	
	ele_params = []
	if len(sorted_pipe_fittings) > 0:
		ele_params = sorted_pipe_fittings[0].Parameters
	else:
		ele_params = sorted_pipe_accessories[0].Parameters
	ele_param_defs = [p.Definition for p in ele_params]
	ele_param_n = [p.Name for p in ele_param_defs]
	#test.append(sorted_elements)
	
	w_no_1_param = ()
	w_no_2_param = ()
	w_no_3_param = ()
	
	number_of_welds = ()
	
	for p in ele_params:
		if p.Definition.Name == "MRT_WeldNo_1":
			w_no_1_param = p.GUID
		if p.Definition.Name == "MRT_WeldNo_2":
			w_no_2_param = p.GUID
		if p.Definition.Name == "MRT_WeldNo_3":
			w_no_3_param = p.GUID
		if p.Definition.Name == "MRT_Automation_NumberOfWelds":
			number_of_welds = p.GUID
	
	
	#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PART 4 - TAG ELEMENTS ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
	
	if run:
	
		trans.Start("Tag Element")
		
		view_locked = spool_views[n].SaveOrientationAndLock()
		
		# --- GET TAG FAMILYSYMBOL TYPES ---
		
		famTypes = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
		
		weld_tag_no1 = ()
		weld_tag_no2 = ()
		weld_tag_no3 = ()
		
		
		for f in famTypes:
			if Element.Name.GetValue(f) == "Weld No. 1":
				weld_tag_no1 = f
			if Element.Name.GetValue(f) == "Weld No. 2":
				weld_tag_no2 = f
			if Element.Name.GetValue(f) == "Weld No. 3":
				weld_tag_no3 = f
			
			
		# --- CREATE WELD TAGS ---
			
		
		weld_nos_1 = []
		weld_nos_2 = []
		weld_nos_3 = []
		weld_nos_1_st = []
		weld_nos_2_st = []
		weld_nos_3_st = []
			
			
		for i in sorted_elements:
			
			p_l = i.Location.Point	
			w_no_1 = i.get_Parameter(w_no_1_param)
			w_no_2 = i.get_Parameter(w_no_2_param)
			w_no_3 = i.get_Parameter(w_no_3_param)
			w_no_1_st = i.get_Parameter(w_no_1_param).AsString()
			w_no_2_st = i.get_Parameter(w_no_2_param).AsString()
			w_no_3_st = i.get_Parameter(w_no_3_param).AsString()
			weld_nos_1.append(w_no_1)
			weld_nos_2.append(w_no_2)
			weld_nos_3.append(w_no_3)
			weld_nos_1_st.append(w_no_1_st)
			weld_nos_2_st.append(w_no_2_st)
			weld_nos_3_st.append(w_no_3_st)
			
			no_weldz = i.get_Parameter(number_of_welds).AsString()
			if no_weldz != None:
				no_welds = int(no_weldz)
			
				if no_welds > 0:
					i.get_Parameter(w_no_1_param).Set(str(numero))
					numero+=1
					tag_1 = IndependentTag.Create(doc, weld_tag_no1.Id, spool_views[n].Id, Reference(i), True, TagOrientation.Horizontal, p_l+XYZ(-50,-50,0))
					tagz.append(tag_1)
					if no_welds > 1:
						i.get_Parameter(w_no_2_param).Set(str(numero))
						numero+=1
						tag_2 = IndependentTag.Create(doc, weld_tag_no2.Id, spool_views[n].Id, Reference(i), True, TagOrientation.Horizontal, p_l+XYZ(0,0,0))
						tag_2.TagHeadPosition += XYZ(0,0,0)
						tagz.append(tag_2)
						if no_welds > 2:
							i.get_Parameter(w_no_3_param).Set(str(numero))
							numero+=1
							tag_3 = IndependentTag.Create(doc, weld_tag_no3.Id, spool_views[n].Id, Reference(i), True, TagOrientation.Horizontal, p_l+XYZ(50,50,0))
							tag_3.TagHeadPosition += XYZ(0,0,0)
							tagz.append(tag_3)
	
	
		trans.Commit()
		



		
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PART 5 - TRANSACTION 2 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# --- CHANGE TAG FAMILY TYPES ---

#TransactionManager.Instance.EnsureInTransaction(doc)	

#tagz[0].ChangeTypeId(pipe_a_tag_types[1].Id)

#param_type.ChangeTypeId(

#param_type.Set(2)



#TransactionManager.Instance.TransactionTaskDone()

	
#OUT= zip_01, sorted, sorted_elements, sorted_locations
#OUT = tagz, param_n, param_type_value, fam_filtered, pipe_f_tag_types, test, pipe_f_weld_tag_types, pipe_a_weld_tag_types
#OUT = a_weld_nos_1, a_weld_nos_2, a_weld_nos_3, f_weld_nos_1, f_weld_nos_2, f_weld_nos_3, a_w_param_n, f_w_no_3_param, a_w_no_3_param, f_w_no_2_param, a_w_no_2_param
#OUT =tagz, numero

#OUT = spool_views, tagz, test
OUT = PAM_no_OG, tagz, PAM_spool_sheets, test, spool_views, views_OG, TESTER