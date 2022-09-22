"PAM frames and sheet"
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



doc = __revit__.ActiveUIDocument.Document
#docTitle = doc.Title
uidoc = __revit__.ActiveUIDocument
#uiapp = DocumentManager.Instance.CurrentUIApplication
#app = uiapp.Application

doc1=__revit__.ActiveUIDocument

app = __revit__.Application

__title__ = "PFS"

PAM_no_OG = 'Enter PAM Number'
Frame_no = 'Enter Frame Number' 


class MyClass(Form):

    def __init__(self):
        global PAM_no_OG
        global Frame_no
        name = 0
        self.Text = 'Text Widget Demo'
        self.label = Label()
        self.label.Text = "This is text widget Demo"
        self.label.Location = Point(100, 150)
        self.label.Height = 50
        self.label.Width = 250
        self.textbox = TextBox()
        self.textbox.Text = str(PAM_no_OG)
        self.textbox.Enabled = True
        self.textbox.Location = Point(50, 50)
        self.textbox.Width = 200
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
        self.textbox2 = TextBox()
        self.textbox2.Text = str(Frame_no)
        self.textbox2.Location = Point(100,100)
        self.textbox2.Width = 200
        self.Controls.Add(self.textbox2)


    
    
    
    
    
    
    def OnClick(self,sender,args):
        MessageBox.Show(self.textbox.Text)
        global PAM_no
        global Frame_no
         
        PAM_no = self.textbox.Text
        Frame_no = self.textbox2.Text
        #Close()
        #def Close(self):
        #self.close()
        #form = LabelDemoForm()
dapp.Run(MyClass())







#PAM_no_OG = IN[0]
#Frame_no = IN[1]
#Frame_no = PAM_nos_string.split(",")

#Assigning the transaction 

trans = Transaction(doc)


#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PART 1 - VIEW SETUP ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# --- MISC DEFINITIONS ---

utid = UnitTypeId.Millimeters

def conv(x):
	c = UnitUtils.ConvertToInternalUnits(x,utid)
	return c
	
def change_name(x,y):
	param_viewname = x.get_Parameter(BuiltInParameter.VIEW_NAME)
	param_viewname.Set(y)
	
def sched_change_name(s,n):
	param_viewname = s.get_Parameter(BuiltInParameter.VIEW_NAME)
	clone_name = s.Name
	new_name_1 = clone_name.replace("Copy 1", " - " + n)
	param_viewname.Set(new_name_1)
	
def prefix(assembly_id):
	el = doc.GetElement(assembly_id)
	n = el.Name
	return n
	
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
	
	trans.Start("View Setup")	
	ElementTransformUtils.RotateElement(doc, cbox.Id, axis, 90*(math.pi/180))	
	doc.Regenerate()

	trans.Commit() 
	return view
	
# --- GET TITLEBLOCKS ---
"""
titleblocks = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_TitleBlocks)

t_names = []
t_search = []
#t_test =t_search[0]

for t in titleblocks:
	UnwrapElement(t)
	t_id = t.Id
	ele = doc.GetElement(t_id).FamilyName
	if ele.Contains("MRT_ANO_A0_TB_2019"):
		t_search.append(t_id)
	t_names.append(ele)
	
"""

project_view_list = FilteredElementCollector(doc).OfClass(View)

project_view_templates = []
view_templates = []

vt_axo = ()
vt_plan = ()
vt_ele = ()

for v in project_view_list:
	if v.IsTemplate == True:
		project_view_templates.append(v)

for t in project_view_templates:
	t_name = t.Name
	if (str(t_name)) == "ASM_PAMFRAME_3D":
		vt_axo = t	
		view_templates.append(t)
	if (str(t_name)) == "ASM_PAMFRAME_PLAN":
		vt_plan = t	
		view_templates.append(t)
	if (str(t_name)) == "ASM_PAMFRAME_ELEVATION":
		vt_ele = t	
		view_templates.append(t)
		
# --- GET ASSEMBLY ---

all_assemblies = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Assemblies)
all_assemblies.WhereElementIsNotElementType()
all_assemblies.ToElements()

PAM_assembly = []

for a in all_assemblies:
	n = a.Name
	if n.Contains(PAM_no_OG):
		PAM_assembly.append(a)

assemblies_id = []
for a in PAM_assembly:
	assemblies_id.append(a.Id)
	
assembly_names = []
for a in PAM_assembly:
	assembly_names.append(a.Name)
	
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PART 2 - VIEW CREATION ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
	
# --- CREATE FRAME VIEWS ---

errors = []
viewlist = ()

trans.Start("View Creation")

if len(PAM_assembly) == 1:

	

	for a in assemblies_id:
		views=[]
		spools=[]
		axo = AssemblyViewUtils.Create3DOrthographic(doc,a,vt_axo.Id,True)
		change_name(axo, str(Frame_no) + " - ISOMETRIC VIEW")
		views.append(axo)	
		
		plan_view = AssemblyViewUtils.CreateDetailSection(doc,a,AssemblyDetailViewOrientation.ElevationTop, vt_plan.Id, True)
		change_name(plan_view, str(Frame_no) +  " - PLAN VIEW")
		cbox = GetViewCropBoxElement(plan_view)
		plan_view_r = RotateCropBox(plan_view, cbox)
		views.append(plan_view_r)	
		
		elevation_right = AssemblyViewUtils.CreateDetailSection(doc,a,AssemblyDetailViewOrientation.ElevationRight, vt_ele.Id, True)
		change_name(elevation_right, str(Frame_no) + " - ELEVATION A")
		views.append(elevation_right)
		
		elevation_front = AssemblyViewUtils.CreateDetailSection(doc,a,AssemblyDetailViewOrientation.ElevationFront, vt_ele.Id, True)
		change_name(elevation_front, str(Frame_no) + " - ELEVATION B")
		views.append(elevation_front)
		
		elevation_back = AssemblyViewUtils.CreateDetailSection(doc,a,AssemblyDetailViewOrientation.ElevationLeft, vt_ele.Id, True)
		change_name(elevation_back, str(Frame_no) + " - ELEVATION C")
		views.append(elevation_back)
		
		viewlist = views
		
		
		
elif len(PAM_assembly) == 0:
	errors.append("ERROR: No Assembly found with number: " + str(PAM_no_OG))
else:
	errors.append("ERROR: ore than one Assembly found with number: " + str(PAM_no_OG))
	
doc.Regenerate()
trans.Commit()

# --- GET SHEET ---

sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets)
sheets.WhereElementIsNotElementType()


sheet = []
template_sheet = ()

for s in sheets:
	sheet_no = s.SheetNumber
	if sheet_no.Contains(str(PAM_no_OG)):
		#unwrapped = UnwrapElement(s)
		sheet.append(s)
	if sheet_no == "AT_2.2":
		#unwrapped = UnwrapElement(s)
		template_sheet = s	
		
def sheet_name(s):
	param_name = s.get_Parameter(BuiltInParameter.SHEET_NAME)
	param_number = s.get_Parameter(BuiltInParameter.SHEET_NUMBER)
	new_number = str(Frame_no)
	param_number.Set(new_number)
	new_name = "PAM FRAME FABRICATION DRAWING"
	param_name.Set(new_name)
	return new_number

# --- GET A1 TITLEBLOCK ---
	
titleblock = FilteredElementCollector(doc, template_sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks)

tblock= []
for t in titleblock:
	#UnwrapElement(t)
	sy = t.Symbol
	s_id = sy.Id
	tblock.append(s_id)


# --- VIEWS ON SHEET ---

trans.Start("View on Sheet")

new_sheet = Autodesk.Revit.DB.ViewSheet.Create(doc, tblock[0])
sheet_name(new_sheet)	
	
	
s_loc_01 = XYZ(conv(585),conv(450),0)
s_loc_02 = XYZ(conv(185),conv(465),0)
s_loc_03 = XYZ(conv(170),conv(270),0)
s_loc_04 = XYZ(conv(360),conv(270),0)	
s_loc_05 = XYZ(conv(170),conv(120),0)	

	

center_points = [s_loc_01,s_loc_02,s_loc_03,s_loc_04,s_loc_05]	
	
for v,c in zip(viewlist,center_points):
	if Viewport.CanAddViewToSheet(doc,new_sheet.Id,v.Id) == True:
		Viewport.Create(doc,new_sheet.Id,v.Id,c)

doc.Regenerate()

trans.Commit() 

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PART 3 - ASSEMBLY SCHEDULES ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# --- VIEW TEMPLATES ---

vt_structuralFraming = ()
vt_structuralColumns = ()

for t in project_view_templates:
	t_name = t.Name
	if (str(t_name)) == "SCH_ASM_StructuralFraming":
		vt_structuralFraming = t
	if (str(t_name)) == "SCH_ASM_StructuralColumns":
		vt_structuralColumns = t
		
# --- CATEGORIES ---

cat_structcolumn = ()
cat_structframing = ()

categories = [i for i in doc.Settings.Categories]
category_name = []

for c in categories:
	cat_name = c.Name
	category_name.append(cat_name)
	
	if cat_name == "Structural Columns":
		cat_structcolumn = c
	if cat_name == "Structural Framing":
		cat_structframing = c	

# --- CREATE SCHEDULES ---

schedule_list=[]		
			
# Start Transaction
trans.Start("Create Schdules")
for p in PAM_assembly:
	a = p.Id
	s1=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_structcolumn.Id,vt_structuralColumns.Id,True)
	schedule_list.append(s1)
	s2=AssemblyViewUtils.CreateSingleCategorySchedule(doc,a,cat_structframing.Id,vt_structuralFraming.Id,True)
	schedule_list.append(s2)
	
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
		
	#change_name(s1,"Rectangular Duct Schedule")
	

doc.Regenerate()
# End Transaction
trans.Commit()

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PART 4 - SCHEDULES ON SHEET ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# --- FORMAT SCHEDULES ---	


trans.Start("Format Schedules")

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
trans.Start("View Placement")

scheds_on_sheet = []

x = conv(478)
y = conv(280)

point_0 = XYZ(x,y,0)

sched_0 = ScheduleSheetInstance.Create(doc,new_sheet.Id,schedule_list[0].Id,point_0)
scheds_on_sheet.append(sched_0)


def get_y(sched):
	bb = sched.get_BoundingBox(new_sheet)
	bb_min = bb.Min
	y = bb_min.Y
	return y

flexi_schedules = schedule_list[1:]

schedule_no = 0

for f in flexi_schedules:
	y_flexi = get_y(scheds_on_sheet[schedule_no])
	flexi_point = XYZ(x,y_flexi,0)
	sched_create = ScheduleSheetInstance.Create(doc,new_sheet.Id,f.Id,flexi_point)
	scheds_on_sheet.append(sched_create)
	schedule_no += 1
	

doc.Regenerate()
# End Transaction
trans.Commit()
trans.Join()
#OUT= PAM_no_OG, PAM_nos_string

if len(errors) == 0:
	errors = "Script ran without errors"
OUT = PAM_no_OG, Frame_no