from tokenize import Double
import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *
# Import ToDSType(bool) extension method
clr.AddReference("RevitNodes")
import Revit
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

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

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction
from Autodesk.Revit.DB import ElementParameterFilter, FilterElementIdRule, ParameterValueProvider, ElementId
from Autodesk.Revit.DB import BuiltInParameter, FilterNumericEquals, ElementTransformUtils
from Autodesk.Revit.DB import XYZ, Viewport, AssemblyViewUtils, ScheduleHorizontalAlignment, ScheduleSheetInstance
from Autodesk.Revit.DB import View
from Autodesk.Revit.DB import AssemblyDetailViewOrientation
from Autodesk.Revit.DB import UnitTypeId, UnitUtils, Document
from Autodesk.Revit.UI import UIApplication
from Autodesk.Revit.ApplicationServices import Application as app_

from System.Windows.Forms import Application as dapp
from System.Windows.Forms import Form, TextBox, Label, Button, MessageBox, CheckBox
from System.Windows.Forms import ToolBar, ToolBarButton, OpenFileDialog, FolderBrowserDialog, ToolStripButton, ToolStrip
from System.Windows.Forms import DialogResult, ScrollBars, DockStyle

from System.Drawing import Point

import math 
from pyrevit import revit, DB,UI

import os
import datetime

import math
import traceback
import csv

import itertools
from itertools import izip_longest

doc = __revit__.ActiveUIDocument.Document
#docTitle = doc.Title
uidoc = __revit__.ActiveUIDocument
#uiapp = DocumentManager.Instance.CurrentUIApplication
#app = uiapp.Application

doc1=__revit__.ActiveUIDocument

app = __revit__.Application


__title__ = "DPAM"
#print("Stops here")

PAM_no = 'PAM Number'
spools_run = False
#_________________________________ GUI ________________________________________________
class MyClass(Form):

	def __init__(self):
		global PAM
		name = 0
		self.Text = 'Text Widget Demo'
		self.label = Label()
		self.label.Text = "This is text widget Demo"
		self.label.Location = Point(100, 150)
		self.label.Height = 50
		self.label.Width = 250
		self.textbox = TextBox()
		self.textbox.Text = str(PAM_no)
		self.textbox.Enabled = True
		self.textbox.Location = Point(50, 50)
		self.textbox.Width = 200
		self.Controls.Add(self.label)
		self.Controls.Add(self.textbox)
		#self.Controls.Add(self.textbox.Text)
		#self.all_c.append(self.textbox.Text)
		PAM = self.textbox.Text
		self.but = Button()
		self.but.Text = 'Ok'
		self.but.Location = Point(75,75)
		self.but.Click += self.OnClick
		self.Controls.Add(self.but)
		#ereturn PAM

		self.CheckBox1 = CheckBox()
		self.CheckBox1.Location = Point(150,100)
		self.CheckBox1.AutoSize = True
		self.CheckBox1.Width = 100
		self.CheckBox1.Text = "Spools Run"
		self.Controls.Add(self.CheckBox1)
	def OnClick(self,sender,args):
		MessageBox.Show(self.textbox.Text)
		global PAM_no
		global spools_run 
		PAM_no = self.textbox.Text
		spools_run = self.CheckBox1.Checked
		#Close()
		#def Close(self):
		#self.close()
		#form = LabelDemoForm()
dapp.Run(MyClass())
		#print(PAM)


#PAM_no = int(PAM_no)




















 

#PAM_no = IN[0]#input from user in a window 

#template_sheet = UnwrapElement(IN[1])
#spools_run = IN[1]#a bool to tick the spools or not

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PART 1 - VIEW SETUP ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
trans= Transaction(doc) #assignign a transaction instance


project_view_list = FilteredElementCollector(doc).OfClass(View)
viewlist=()
spoollist = ()

categories = "HVAC_supply", "HVAC_return", "E", "P"

project_view_templates = []
view_templates = []

vt_axo = ()
vt_hvac_s = ()
vt_hvac_s_plan = ()
vt_hvac_r = ()
vt_hvac_r_plan = ()
vt_e = ()
vt_e_plan = ()
vt_p = ()
vt_p_plan = ()

vt_sys_CA = ()
vt_sys_N2 = ()
vt_sys_CO2 = ()
vt_sys_O2 = ()

for v in project_view_list:
	if v.IsTemplate == True:
		project_view_templates.append(v)

for t in project_view_templates:
	t_name = t.Name
	if (str(t_name)) == "ASM_Z_3D":
		vt_axo = t	
		view_templates.append(t)		
	if (str(t_name)) == "ASM_HVAC_SUPPLY":
		vt_hvac_s = t
		view_templates.append(t)
	if (str(t_name)) == "ASM_HVAC_SUPPLY_PLAN":
		vt_hvac_s_plan = t
		view_templates.append(t)
	if (str(t_name)) == "ASM_HVAC_RETURN":
		vt_hvac_r = t
		view_templates.append(t)
	if (str(t_name)) == "ASM_HVAC_RETURN_PLAN":
		vt_hvac_r_plan = t	
		view_templates.append(t)
	if (str(t_name)) == "ASM_E_ELEVATION":
		vt_e = t
		view_templates.append(t)
	if (str(t_name)) == "ASM_E_PLAN":
		vt_e_plan = t
		view_templates.append(t)		
	if (str(t_name)) == "ASM_P_ELEVATION/SECTION":
		vt_p = t
		view_templates.append(t)
	if (str(t_name)) == "ASM_P_PLAN":
		vt_p_plan = t
		view_templates.append(t)
		
	if (str(t_name)) == "ASM_Z_SYSTEM_CompAir" or (str(t_name)) == "ASM_Z_SYSTEM_CA":
		vt_sys_CA = t
		view_templates.append(t)
	if (str(t_name)) == "ASM_Z_SYSTEM_N2":
		vt_sys_N2 = t
		view_templates.append(t)
	if (str(t_name)) == "ASM_Z_SYSTEM_CO2":
		vt_sys_CO2 = t
		view_templates.append(t)
	if (str(t_name)) == "ASM_Z_SYSTEM_O2":
		vt_sys_O2 = t
		view_templates.append(t)		
				
		
def get_vt(cat,type):
	vt = ()
	if cat == "HVAC_supply":
		if type == "plan":
			vt = vt_hvac_s_plan
		else:
			vt = vt_hvac_s
	if cat == "HVAC_return":
		if type == "plan":
			vt = vt_hvac_r_plan
		else:
			vt = vt_hvac_r		
	if cat == "E":
		if type == "plan":
			vt = vt_e_plan
		else:
			vt = vt_e
	if cat == "P":
		if type == "plan":
			vt = vt_p_plan
		else:
			vt = vt_p
	return vt

def change_name(x,y):
	param_viewname = x.get_Parameter(BuiltInParameter.VIEW_NAME)
	param_viewname.Set(y)
	
def view_name(cat,type):
	name = ()
	if cat == "HVAC_supply":
		if type == "top":
			name = "SUPPLY PLAN"
		if type == "right":
			name = "ELEVATION 2A"
		if type == "front":
			name = "ELEVATION 2B"
	if cat == "HVAC_return":
		if type == "top":
			name = "EXTRACT PLAN"
		if type == "right":
			name = "ELEVATION 3A"
		if type == "front":
			name = "ELEVATION 3B"
	if cat == "E":
		if type == "top":
			name = "ELECTRICAL PLAN"
		if type == "right":
			name = "ELEVATION 4A"
		if type == "front":
			name = "ELEVATION 4B"
	if cat == "P":
		if type == "top":
			name = "PIPEWORK PLAN"
		if type == "right":
			name = "ELEVATION 5A"
		if type == "front":
			name = "ELEVATION 5B"
	return name

# --- ROTATE CBOX DEF ---

def GetViewCropBoxElement(view):
	collector = FilteredElementCollector(doc).WherePasses(ElementParameterFilter(FilterElementIdRule(ParameterValueProvider(ElementId(BuiltInParameter.ID_PARAM)),FilterNumericEquals(), view.Id)))
	collector = [i for i in collector]
	return collector[0]

def RotateCropBox(view, cropBox):
	box = view.CropBox
	transform = box.Transform
	min = transform.OfPoint(box.Min)
	max = transform.OfPoint(box.Max)
	loca = (min+max)/2
	axis = Line.CreateBound(loca,XYZ(loca.X,loca.Y,loca.Z+1))
	
	#t.Start("Rotate CBOX DEF")	
	ElementTransformUtils.RotateElement(doc, cbox.Id, axis, 90*(math.pi/180))	
	doc.Regenerate()

	#t.Commit() 
	return view

# --- GET ASSEMBLY INFO ---

all_assemblies = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Assemblies)
all_assemblies.WhereElementIsNotElementType()
all_assemblies.ToElements()

PAM_assembly = []

for a in all_assemblies:
	n = a.Name
	if n.Contains(PAM_no):
		PAM_assembly.append(a)

assemblies_id = []
for a in PAM_assembly:
	assemblies_id.append(a.Id)
	
assembly_names = []
for a in PAM_assembly:
	assembly_names.append(a.Name)


def prefix(assembly_id):
	el = doc.GetElement(assembly_id)
	n = el.Name
	return n

# --- CREATE ASSEMBLY VIEWS ---
test = []

trans.Start('Create Assembly Views')

#trans1.Start()

for a in assemblies_id:
	views=[]
	spools=[]
	axo = AssemblyViewUtils.Create3DOrthographic(doc,a,vt_axo.Id,True)
	change_name(axo, prefix(a) + " - ISOMETRIC VIEW")
	views.append(axo)
	
	#number_spools = 4
	#for n in range (0,number_spools):
	
	spool_axo_CA = AssemblyViewUtils.Create3DOrthographic(doc,a,vt_sys_CA.Id,True)
	change_name(spool_axo_CA, prefix(a) + " - " + "CA")
	spool_axo_N2 = AssemblyViewUtils.Create3DOrthographic(doc,a,vt_sys_N2.Id,True)
	change_name(spool_axo_N2, prefix(a) + " - " + "N2")
	spool_axo_CO2 = AssemblyViewUtils.Create3DOrthographic(doc,a,vt_sys_CO2.Id,True)
	change_name(spool_axo_CO2, prefix(a) + " - " + "CO2")
	spool_axo_O2 = AssemblyViewUtils.Create3DOrthographic(doc,a,vt_sys_O2.Id,True)
	change_name(spool_axo_O2, prefix(a) + " - " + "O2")
	spools.append(spool_axo_CA)
	spools.append(spool_axo_N2)
	spools.append(spool_axo_CO2)
	spools.append(spool_axo_O2)
	
	for c in categories:
		plan_view = AssemblyViewUtils.CreateDetailSection(doc,a,AssemblyDetailViewOrientation.ElevationTop, (get_vt(c,"plan")).Id, True)
		change_name(plan_view, prefix(a) + " - " + view_name(c,"top"))
		cbox = GetViewCropBoxElement(plan_view)
		plan_view_r = RotateCropBox(plan_view, cbox)
		#TransactionManager.Instance.EnsureInTransaction(doc)
		views.append(plan_view_r)		
		elevation_right = AssemblyViewUtils.CreateDetailSection(doc,a,AssemblyDetailViewOrientation.ElevationRight, (get_vt(c,"ele")).Id, True)
		change_name(elevation_right,prefix(a) + " - " + view_name(c,"right"))
		views.append(elevation_right)		
		elevation_front = AssemblyViewUtils.CreateDetailSection(doc,a,AssemblyDetailViewOrientation.ElevationFront, (get_vt(c,"ele")).Id, True)
		change_name(elevation_front,prefix(a) + " - " + view_name(c,"front"))
		views.append(elevation_front)	
	viewlist = views
	spoollist = spools


#Test = Viewport.CanAddViewToSheet(doc,sheet.Id,reference_viewports[3].Id)

doc.Regenerate()
# End Transaction
trans.Commit()

#trans1.Commit()


#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PART 2 - VIEW PLACEMENT ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#




# --- MISC DEFINITIONS ---

utid = UnitTypeId.Millimeters

def conv(x):
	c = UnitUtils.ConvertToInternalUnits(x,utid)
	return c
		
# --- GET SHEETs ---

sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets)
sheets.WhereElementIsNotElementType()


# sheet_names = []
sheet_no_list = []
sheet = ()
spool_template = ()
template_sheet = ()

for s in sheets:
	sheet_no = s.SheetNumber
	if sheet_no.Contains(PAM_no):
		#unwrapped = UnwrapElement(s)
		sheet = s	
	if sheet_no == "AT_2.0":
		#unwrapped = UnwrapElement(s)
		template_sheet = s	
	if sheet_no == "AT_2.2":
		#unwrapped = UnwrapElement(s)
		spool_template = s	
		
		
# --- GET TITLEBLOCK ---

titleblock = FilteredElementCollector(doc, spool_template.Id).OfCategory(BuiltInCategory.OST_TitleBlocks)

tblock= []
for t in titleblock:
	#UnwrapElement(t)
	sy = t.Symbol
	s_id = sy.Id
	tblock.append(s_id)



# --- NEW SPOOL SHEET --

spool_sheets = []

def sheet_name(n,s,no):
	param_name = n.get_Parameter(BuiltInParameter.SHEET_NAME)
	param_number = n.get_Parameter(BuiltInParameter.SHEET_NUMBER)
	number_OG = s.SheetNumber
	new_number = str(number_OG) + "_" + "0" + str(no)
	param_number.Set(new_number)
	new_name = str(s.Name)
	param_name.Set(new_name)
	return new_number
	
# --- GET COORDINATES ---

viewport_ids = template_sheet.GetAllViewports()
reference_viewports = []
viewport_names = []

for i in viewport_ids:
	v = doc.GetElement(i)
	v_id = v.ViewId
	v_el = doc.GetElement(v_id)
	v_name = v_el.Name
	reference_viewports.append(v)
	viewport_names.append(v_name)

center_points = []

for v in reference_viewports:
	c = v.GetBoxCenter()
	center_points.append(c)
	
v_zip = list(zip(viewport_names,center_points))

# --- SPOOL LOCATIONS --- 

s_loc_01 = XYZ(conv(150),conv(520),0)
s_loc_02 = XYZ(conv(450),conv(520),0)
s_loc_03 = XYZ(conv(150),conv(220),0)
s_loc_04 = XYZ(conv(450),conv(220),0)

spool_locate = [s_loc_01,s_loc_02,s_loc_03,s_loc_04]

# --- ORDER VIEWLIST ---

sortedViews = ()


if str(template_sheet.SheetNumber).Contains("AT_2.0"):
	sortedIndex = [1,4,0,2,3,5,6,7,10,8,9,11,12]
	sortedViews = [viewlist[i] for i in sortedIndex]
	
sortedViewIds = []

for i in sortedViews:
	iD = i.Id
	sortedViewIds.append(iD)

# --- SHEET PARAMETERS ---

sheet_params = sheet.Parameters
	
sheet_grp = ()
sheet_subgrp = ()

for p in sheet_params:
	if p.Definition.Name == "MRT_BRO_Sheet Grouping":
		sheet_grp = p.GUID
	if p.Definition.Name == "MRT_BRO_Sheet Sub Grouping":
		sheet_subgrp = p.GUID

# --- VIEWPORT PLACEMENT ---

def elementId(x):
	ei = Autodesk.Revit.DB.ElementId(x)
	return ei

# Start Transaction
trans.Start('Viewport Placement')

views_on_sheet = []

#if input_bool_placesheets == 1:

for v,c in zip(sortedViews,center_points):
	if Viewport.CanAddViewToSheet(doc,sheet.Id,v.Id) == True:
		Viewport.Create(doc,sheet.Id,v.Id,c)
		

if spools_run == True:
	
	new_spool_sheet = Autodesk.Revit.DB.ViewSheet.Create(doc, tblock[0])
	spool_sheet = sheet_name(new_spool_sheet,sheet, 1)
	sheet_g = new_spool_sheet.get_Parameter(sheet_grp).Set("MECHANICAL")
	sheet_subg = new_spool_sheet.get_Parameter(sheet_subgrp).Set("FABRICATION")
	for n in range(1, (len(spoollist))+1):
		Viewport.Create(doc,new_spool_sheet.Id,(spoollist[n-1]).Id,spool_locate[n-1])
		

vps = sheet.GetAllViewports()
viewport_params = doc.GetElement(vps[0]).Parameters
v_params = [i.Definition.Name for i in viewport_params]

doc.Regenerate()
# End Transaction
trans.Commit() 

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PART 3 - ASSEMBLY SCHEDULES ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# --- VIEW TEMPLATES ---

vt_ductrect = ()
vt_ductrectshoe = ()
vt_ductrectbend = ()
vt_ductcircular = ()
vt_ductfittings = ()
vt_ductaccessories = ()
vt_pipes = ()
vt_pipefittings = ()
vt_pipeaccessories = ()
vt_pipeclamps = ()
vt_cabletray = ()
vt_busbar = ()
vt_busbarendfeed = ()
vt_busbartapoffs= ()

for t in project_view_templates:
	t_name = t.Name
	if (str(t_name)) == "SCH_ASM_Pipe":
		vt_pipes = t
	if (str(t_name)) == "SCH_ASM_PipeFittings":
		vt_pipefittings = t	
	if (str(t_name)) == "SCH_ASM_PipeAccessories":
		vt_pipeaccessories = t
	if (str(t_name)) == "SCH_ASM_Pipe_Clamps":
		vt_pipeclamps = t	
	if (str(t_name)) == "SCH_ASM_DuctRect":
		vt_ductrect = t
	if (str(t_name)) == "SCH_ASM_DuctRectShoe":
		vt_ductrectshoe = t			
	if (str(t_name)) == "SCH_ASM_DuctRectBend":
		vt_ductrectbend = t	
	if (str(t_name)) == "SCH_ASM_DuctCircular":
		vt_ductcircular = t	
	if (str(t_name)) == "SCH_ASM_DuctFittings":
		vt_ductfittings = t	
	if (str(t_name)) == "SCH_ASM_DuctAccessories":
		vt_ductaccessories = t	
	if (str(t_name)) == "SCH_ASM_CableTray":
		vt_cabletray = t
	if (str(t_name)) == "SCH_ASM_Busbar":
		vt_busbar = t		
	if (str(t_name)) == "SCH_ASM_BusbarEndFeed":
		vt_busbarendfeed = t
	if (str(t_name)) == "SCH_ASM_BusbarTapOffs":
		vt_busbartapoffs = t	
		
# --- CATEGORIES ---

cat_ducts = ()
cat_ductfittings = ()
cat_ductaccessories = ()
cat_pipes = ()
cat_pipefittings = ()
cat_pipeaccessories = ()
cat_cabletrays = ()
cat_cabletrayfittings = ()
cat_electricalequipment = ()

categories = [i for i in doc.Settings.Categories]
category_name = []
category_search = []

for c in categories:
	cat_name = c.Name
	category_name.append(cat_name)
	if cat_name.Contains("Electric"):
		category_search.append(cat_name)
	if cat_name == "Ducts":
		cat_ducts = c
	if cat_name == "Duct Fittings":
		cat_ductfittings = c
	if cat_name == "Duct Accessories":
		cat_ductaccessories = c
	if cat_name == "Pipes":
		cat_pipes = c	
	if cat_name == "Pipe Fittings":
		cat_pipefittings = c
	if cat_name == "Pipe Accessories":
		cat_pipeaccessories = c
	if cat_name == "Cable Trays":
		cat_cabletrays = c
	if cat_name == "Cable Tray Fittings":
		cat_cabletrayfittings = c
	if cat_name == "Electrical Equipment":
		cat_electricalequipment = c

# --- CREATE SCHEDULES ---

schedule_list=[]		
			
# Start Transaction
trans.Start('Create Schedules')
for p in PAM_assembly:	
	a = p.Id
	s1=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_ducts.Id,vt_ductrect.Id,True)
	schedule_list.append(s1)
	s2=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_ductfittings.Id,vt_ductrectshoe.Id,True)
	schedule_list.append(s2)
	s3=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_ductfittings.Id,vt_ductrectbend.Id,True)
	schedule_list.append(s3)
	s4=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_ducts.Id,vt_ductcircular.Id,True)
	schedule_list.append(s4)
	s5=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_ductfittings.Id,vt_ductfittings.Id,True)
	schedule_list.append(s5)
	s6=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_ductaccessories.Id,vt_ductaccessories.Id,True)
	schedule_list.append(s6)
	s7=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_pipes.Id, vt_pipes.Id,True)
	schedule_list.append(s7)
	s8=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_pipefittings.Id, vt_pipefittings.Id,True)
	schedule_list.append(s8)
	s9=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_pipeaccessories.Id, vt_pipeaccessories.Id,True)
	schedule_list.append(s9)
	s10=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_pipeaccessories.Id, vt_pipeclamps.Id,True)
	schedule_list.append(s10)
	s11=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_cabletrays.Id, vt_cabletray.Id,True)
	schedule_list.append(s11)
	s12=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_cabletrayfittings.Id, vt_busbar.Id,True)
	schedule_list.append(s12)
	s13=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_cabletrayfittings.Id, vt_busbarendfeed.Id,True)
	schedule_list.append(s13)
	s14=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_electricalequipment.Id, vt_busbartapoffs.Id,True)
	schedule_list.append(s14)
	
	params = []
	
	for p in s1.Parameters:
		pn = p.Definition.Name
		params.append(pn)
		try:
			params.append(p.GUID)
		except:
			params.append(p.Definition.BuiltInParameter)
		params.append("-"*20)
		
	def change_name(x,y):
		param_viewname = x.get_Parameter(BuiltInParameter.VIEW_NAME)
		param_viewname.Set(y)
		
	change_name(s1,"Rectangular Duct Schedule")
	change_name(s2,"Rectangular Shoe Schedule")
	change_name(s3,"Rectangular Duct Bend Schedule")
	change_name(s4,"Circular Duct Schedule")
	change_name(s5,"Duct Fitting Schedule")
	change_name(s6,"Duct Accessory Schedule")
	change_name(s7,"Pipe Schedule")
	change_name(s8,"Pipe Fitting Schedule")
	change_name(s9,"Pipe Accessory Schedule")
	change_name(s10,"Pipe Clamp Schedule")
	change_name(s11,"Cable Tray Schedule")
	change_name(s12,"Busbar Schedule")
	change_name(s13,"Busbar End Feed Schedule")
	change_name(s14,"Busbar Tap Offs Schedule")

doc.Regenerate()
# End Transaction
trans.Commit()

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PART 4 - SCHEDULES ON SHEET ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# --- FORMAT SCHEDULES ---	


trans.Start('Format Schedules')

def getField(schedule, name):
	definition = schedule.Definition
	count = definition.GetFieldCount()
	for i in range(0, count, 1):
		if definition.GetField(i).GetName() == name:
			field = definition.GetField(i)
	return field
	
def getField_list(schedule):
	definition = schedule.Definition
	count = definition.GetFieldCount()
	fields = []
	for i in range(0, count, 1):
		field = definition.GetField(i)
		fields.append(field)
	return fields

for p in schedule_list:
	definition = p.Definition
	count = definition.GetFieldCount()
	fields = []
	for i in range(0, count, 1):
		field = definition.GetField(i)
		if field.IsHidden != True:
			fields.append(field)
	field_width = conv(210/len(fields))
	for f in fields:
		sc = ScheduleHorizontalAlignment.Center
		f.HorizontalAlignment = sc
		f.SheetColumnWidth = field_width


doc.Regenerate()
trans.Commit() 


# --- VIEWPORT PLACEMENT ---

# Start Transaction
trans.Start('Viewport Placement')

scheds_on_sheet = []

x = conv(850)
y = conv(550)

point_0 = XYZ(x,y,0)

sched_0 = ScheduleSheetInstance.Create(doc,sheet.Id,schedule_list[0].Id,point_0)
scheds_on_sheet.append(sched_0)


def get_y(sched):
	bb = sched.get_BoundingBox(sheet)
	bb_min = bb.Min
	y = bb_min.Y
	return y

flexi_schedules = schedule_list[1:]

schedule_no = 0

while schedule_no <= 3:
	for f in flexi_schedules:
		y_flexi = get_y(scheds_on_sheet[schedule_no])
		flexi_point = XYZ(x,y_flexi,0)
		sched_create = ScheduleSheetInstance.Create(doc,sheet.Id,f.Id,flexi_point)
		scheds_on_sheet.append(sched_create)
		schedule_no += 1
	

doc.Regenerate()
# End Transaction
trans.Commit()

#OUT=viewlist, assembly_names, spoollist
#OUT = assembly_names[0], spools_run, v_params
#OUT = PAM_no