"Takes the desired direcotry for export and generate a BOX CSV from all avaiable documents in the project file"
import clr 
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction
from Autodesk.Revit.DB import UnitTypeId, UnitUtils, Document
from Autodesk.Revit.UI import UIApplication
from Autodesk.Revit.ApplicationServices import Application as app_
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Drawing import Point
from System.Windows.Forms import Application as dapp
from System.Windows.Forms import Form, TextBox, Label
from System.Windows.Forms import ToolBar, ToolBarButton, OpenFileDialog, FolderBrowserDialog, ToolStripButton, ToolStrip
from System.Windows.Forms import DialogResult, ScrollBars, DockStyle

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *
from pyrevit import revit, DB,UI
import os
import datetime

import math
import traceback
import csv

import itertools
from itertools import izip_longest

__title__ = "BOQ"

class LabelDemoForm(Form):

    def __init__(self):
        self.Text = 'Text Widget Demo'

        self.label = Label()
        self.label.Text = "This is text widget Demo"
        self.label.Location = Point(100, 150)
        self.label.Height = 50
        self.label.Width = 250
        
        self.textbox = TextBox()
        self.textbox.Text = "Baby Text Widget"
        self.textbox.Location = Point(50, 50)
        self.textbox.Width = 200

        self.Controls.Add(self.label)
        self.Controls.Add(self.textbox)

form = LabelDemoForm()
#dapp.Run(form)



path = " "

class IForm(Form):

    def __init__(self):
        self.Text = "Choose Directory"

        toolbar = ToolBar()
        toolbar.Parent = self
        openb = ToolBarButton()
        toolstrip1 = ToolStrip()

        self.textbox = TextBox()
        self.textbox.Parent = self
        self.textbox.Multiline = True
        self.textbox.ScrollBars = ScrollBars.Both
        self.textbox.WordWrap = False
        #self.textbox.Parent = self
        self.textbox.Dock = DockStyle.Fill
        self.textbox.Text = "Directory"
        self.textbox.Width = 100

        toolbar.Buttons.Add(openb)
        toolbar.ButtonClick += self.OnClicked
        #toolstrip1.Items.AddRange[openb]
        #openb.Click +=self.OnClicked
        openb.Text = "Open"

        self.CenterToScreen()

    def OnClicked(self, sender, event):
        dialog = FolderBrowserDialog()
        #dialog.Filter = "C# files (*.cs)|*.cs"
		
        if dialog.ShowDialog(self) == DialogResult.OK:
			global path
			path = dialog.SelectedPath
			#print(path)
			#print(type(path))
			return path


#
dapp.Run(IForm())




utid = UnitTypeId.Meters

def conv(x):
	c = UnitUtils.ConvertFromInternalUnits(x,utid)
	return c

directory =path
dir_replaced = "/"
dir_replacement = "//"
directory_raw = directory.replace(dir_replaced, dir_replacement)




dateStamp   = datetime.datetime.today().strftime("%y%m%d")
timeStamp   = datetime.datetime.today().strftime("%H%M%S")





doc = __revit__.ActiveUIDocument.Document
#docTitle = doc.Title
uidoc = __revit__.ActiveUIDocument
#uiapp = DocumentManager.Instance.CurrentUIApplication
#app = uiapp.Application



doc1=__revit__.ActiveUIDocument

app = __revit__.Application 

#print(app)





class MERIT:

    def __init__(self, Doc):
        self.Doc = Doc

    def Ceilings(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements()
        MRT_Ceiling = [i for i in filteredCollector]
        return MRT_Ceiling

    def CurtainPanels(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_CurtainWallPanels).WhereElementIsNotElementType().ToElements()
        MRT_CurtainPanels = [i for i in filteredCollector]
        return MRT_CurtainPanels

    def CaseWork(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_Casework).WhereElementIsNotElementType().ToElements()
        MRT_Casework = [i for i in filteredCollector]
        return MRT_Casework

    def CurtainWallMullions(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_CurtainWallMullions).WhereElementIsNotElementType().ToElements()
        MRT_CurtainWallMullions = [i for i in filteredCollector]
        return MRT_CurtainWallMullions

    def Doors(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements()
        MRT_Doors = [i for i in filteredCollector]
        return MRT_Doors

    def ElectricalEquipment(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
        MRT_ElectricalEquipment = [i for i in filteredCollector]
        return MRT_ElectricalEquipment

    def Floors(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements()
        MRT_Floors = [i for i in filteredCollector]
        return MRT_Floors

    def GenericModels(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
        MRT_GenericModels = [i for i in filteredCollector]
        return MRT_GenericModels

    def MechanicalEquipment(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()
        MRT_MechanicalEquipment = [i for i in filteredCollector]
        return MRT_MechanicalEquipment

    def Pipes(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
        MRT_Pipes = [i for i in filteredCollector]
        return MRT_Pipes

    def Parking(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_Parking).WhereElementIsNotElementType().ToElements()
        MRT_Parking = [i for i in filteredCollector]
        return MRT_Parking

    def PipeFittings(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements()
        MRT_PipeFittings = [i for i in filteredCollector]
        return MRT_PipeFittings

    def PipeAcessory(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()
        MRT_PipeAcessory = [i for i in filteredCollector]
        return MRT_PipeAcessory
    
    def PipeInsulation(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_PipeInsulations).WhereElementIsNotElementType().ToElements()
        MRT_PipeInsulation = [i for i in filteredCollector]
        return MRT_PipeInsulation

    def PlumbingFixtures(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_PlumbingFixtures).WhereElementIsNotElementType().ToElements()
        MRT_PlumbingFixtures = [i for i in filteredCollector]
        return MRT_PlumbingFixtures

    def Railings(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_Railings).WhereElementIsNotElementType().ToElements()
        MRT_Railings = [i for i in filteredCollector]
        return MRT_Railings

    def Roofs(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToElements()
        MRT_Roofs = [i for i in filteredCollector]
        return MRT_Roofs

    def SecurityDevices(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_SecurityDevices).WhereElementIsNotElementType().ToElements()
        MRT_SecurityDevices = [i for i in filteredCollector]
        return MRT_SecurityDevices

    def Site(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_Site).WhereElementIsNotElementType().ToElements()
        MRT_Site = [i for i in filteredCollector]
        return MRT_Site

    def EdgeSlab(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_EdgeSlab).WhereElementIsNotElementType().ToElements()
        MRT_EdgeSlab = [i for i in filteredCollector]
        return MRT_EdgeSlab

    def SpecialityEquipment(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_SpecialityEquipment).WhereElementIsNotElementType().ToElements()
        MRT_SpecialityEquipment = [i for i in filteredCollector]
        return MRT_SpecialityEquipment

    def Stairs(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_Stairs).WhereElementIsNotElementType().ToElements()
        MRT_Stairs = [i for i in filteredCollector]
        return MRT_Stairs

    def StructuralColumns(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType().ToElements()
        MRT_StructuralColumns = [i for i in filteredCollector]
        return MRT_StructuralColumns

    def StructuralFraming(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().ToElements()
        MRT_StructuralFraming = [i for i in filteredCollector]
        return MRT_StructuralFraming

    def TopRails(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_RailingTopRail).WhereElementIsNotElementType().ToElements()
        MRT_TopRails = [i for i in filteredCollector]
        return MRT_TopRails

    def Walls(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
        MRT_Walls = [i for i in filteredCollector]
        return MRT_Walls

    def Windows(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements()
        MRT_Windows = [i for i in filteredCollector]
        return MRT_Windows

    def HandRails(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_RailingHandRail).WhereElementIsNotElementType().ToElements()
        MRT_HandRails = [i for i in filteredCollector]
        return MRT_HandRails

    def DuctFittings(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()
        MRT_DuctFittings = [i for i in filteredCollector]
        return MRT_DuctFittings

    def Fascias(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_Fascia).WhereElementIsNotElementType().ToElements()
        MRT_Fascia = [i for i in filteredCollector]
        return MRT_Fascia

    def StructuralConnections(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_StructConnections).WhereElementIsNotElementType().ToElements()
        MRT_StructuralConnections = [i for i in filteredCollector]
        return MRT_StructuralConnections

    def StructuralFoundations(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_StructuralFoundation).WhereElementIsNotElementType().ToElements()
        MRT_StructuralFoundations = [i for i in filteredCollector]
        return MRT_StructuralFoundations

    def StructuralFraming(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().ToElements()
        MRT_StructuralFraming = [i for i in filteredCollector]
        return MRT_StructuralFraming

    def StructuralRebar(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
        MRT_Rebar = [i for i in filteredCollector]
        return MRT_Rebar

    def Roofs(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToElements()
        MRT_Roofs = [i for i in filteredCollector]
        return MRT_Roofs

    def AirTerminal(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()
        MRT_AirTerminal = [i for i in filteredCollector]
        return MRT_AirTerminal

    def Ducts(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()
        MRT_Ducts = [i for i in filteredCollector]
        return MRT_Ducts

    def FlexDucts(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsNotElementType().ToElements()
        MRT_FlexDucts = [i for i in filteredCollector]
        return MRT_FlexDucts
    
    def DuctAcessories(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()
        MRT_DuctAcessory = [i for i in filteredCollector]
        return MRT_DuctAcessory

    def DuctFittings(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()
        MRT_FlexFittings = [i for i in filteredCollector]
        return MRT_FlexFittings

    def DuctInsulations(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_DuctInsulations).WhereElementIsNotElementType().ToElements()
        MRT_DuctInsulations = [i for i in filteredCollector]
        return MRT_DuctInsulations

    def CableTray(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().ToElements()
        MRT_CableTray = [i for i in filteredCollector]
        return MRT_CableTray
    
    def CableTrayFittings(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_CableTrayFitting).WhereElementIsNotElementType().ToElements()
        MRT_CableTrayFittings = [i for i in filteredCollector]
        return MRT_CableTrayFittings
    
    def DataDevices(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_DataDevices).WhereElementIsNotElementType().ToElements()
        MRT_DataDevice = [i for i in filteredCollector]
        return MRT_DataDevice

    def LightingFixture(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_LightingDevices).WhereElementIsNotElementType().ToElements()
        MRT_LightingFixture = [i for i in filteredCollector]
        return MRT_LightingFixture

    def FireAlarmDevice(self,):
        filteredCollector = FilteredElementCollector(self.Doc).OfCategory(BuiltInCategory.OST_FireAlarmDevices).WhereElementIsNotElementType().ToElements()
        MRT_FireAlarmDevice = [i for i in filteredCollector]
        return MRT_FireAlarmDevice

#doc = DocumentManager.Instance.CurrentDBDocument
#docTitle = doc.Title
#uiapp = DocumentManager.Instance.CurrentUIApplication
#app = uiapp.Application
#uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ GET REVIT LINKS ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#rvt_links = FilteredElementCollector(doc).OfClass(clr.GetClrType(RevitLinkInstance)).WhereElementIsNotElementType().ToElements()
#rvt_link_ids = FilteredElementCollector(doc).OfClass(clr.GetClrType(RevitLinkInstance)).ToElementIds()

documentz = app.Documents
docNames = [i.Title for i in documentz]

mArc = None
mCiv = None
mStr = None

MRTSYSTypeMarkVALUE = []
MRTSYSDrgNoVALUE = []
MRTSYSystemNoVALUE = []
MRTSYSItemNoVALUE = []

MRTSAAssemblyRefVALUE = []
MRTSAItemNoVALUE = []
MRTSAProdRefVALUE = [] 
MRTSASystemTypeVALUE = []

sizeVALUE = []
assemblyNameVALUE = []
lengthVALUE = []

elementId = []
elementCategory = []
elementGUID = []
elementTYPE = []
elementFAM = []

manufacturerVALUE = []
modelVALUE = []
descriptionVALUE = []
nameVALUE = []
unitVALUE = []
modelName = []



unitsMap = []

em1 = "Nan"
TEST = []

TransferLog = []

for i in documentz:
	if i.Title.Contains("-A-"):
		mArc = i
	elif i.Title.Contains("-C-"):
		mCiv = i
	elif i.Title.Contains("-S-"):
		mStr = i
	else:
		continue

#MRT_ARC = MERIT(mArc)
#MRT_CIV = MERIT(mCiv)
#MRT_STR = MERIT(mStr)
#MRT_DOC = MERIT(doc)
SHT = [doc, mArc, mCiv, mStr]
# sheets = [MRT_DOC, MRT_ARC, MRT_CIV, MRT_STR]
for sheets in SHT:
    try:
        elementsMap = []
        DOCUMENTS = []
        DOCUMENTS.append(MERIT(sheets))
        ComponentsDict = []
        for documents in DOCUMENTS:   
                ME = {'Elements': documents.MechanicalEquipment(), 'Units' : 'EA'} 
                EE = {'Elements': documents.ElectricalEquipment(), 'Units' : 'EA'} 
                D = {'Elements': documents.Ducts(), 'Units' : 'M'} 
                DA = {'Elements': documents.DuctAcessories(), 'Units' : 'EA'}  
                DF = {'Elements': documents.DuctFittings(), 'Units' : 'EA'}  
                FD = {'Elements': documents.FlexDucts(), 'Units' : 'EA'}  
                DI = {'Elements': documents.DuctInsulations(), 'Units' : 'EA'}
                AT = {'Elements': documents.AirTerminal(), 'Units' : 'EA'}
                P = {'Elements': documents.Pipes(), 'Units' : 'EA'}
                PA = {'Elements': documents.PipeAcessory(), 'Units' : 'EA'}
                PF = {'Elements': documents.PipeFittings(), 'Units' : 'EA'}
                PI = {'Elements': documents.PipeInsulation(), 'Units' : ''}
                SC = {'Elements': documents.StructuralColumns(), 'Units' : 'M'}
                SF = {'Elements': documents.StructuralFraming(), 'Units' : 'M'}
                CT = {'Elements': documents.CableTray(), 'Units' : 'EA'}
                CTF = {'Elements': documents.CableTrayFittings(), 'Units' : 'EA'}
                SD = {'Elements': documents.SecurityDevices(), 'Units' : 'EA'}
                DD = {'Elements': documents.DataDevices(), 'Units' : 'EA'}
                LF = {'Elements': documents.LightingFixture(), 'Units' : 'EA'}
                FAD = {'Elements': documents.FireAlarmDevice(), 'Units' : 'EA'}
                RF = {'Elements': documents.Roofs(), 'Units' : 'M2'}
                FL = {'Elements': documents.Floors, 'Units' : 'M2'}
                WL = {'Elements': documents.Walls(), 'Units' : 'M2'}
                CL = {'Elements': documents.Ceilings(), 'Units' : 'M2'}
                CP = {'Elements': documents.CurtainPanels(), 'Units' : 'M2'}
                GM = {'Elements': documents.GenericModels(), 'Units' : 'EA'}
                P = {'Elements': documents.Parking(), 'Units' : 'EA'}
                RL = {'Elements': documents.Railings(), 'Units' : 'EA'}
                S = {'Elements': documents.Site(), 'Units' : 'M2'}
                ST = {'Elements': documents.Stairs(), 'Units' : 'M2'}
                SCN = {'Elements': documents.StructuralConnections(), 'Units' : 'EA'}
                SFD = {'Elements': documents.StructuralFoundations(), 'Units' : 'M3'}
                TR = {'Elements': documents.TopRails(), 'Units' : 'EA'}
                DR = {'Elements': documents.Doors(), 'Units' : 'EA'}
                FCS = {'Elements': documents.Fascias(), 'Units' : 'EA'}
                ES = {'Elements': documents.EdgeSlab(), 'Units' : 'M'}
                SEQ = {'Elements': documents.SpecialityEquipment(), 'Units' : 'EA'}    
                WIN = {'Elements': documents.Windows(), 'Units' : 'EA'}
                
        ComponentsDict = [ME, EE, D, DA, DF, FD, DI, AT, P, PA, PF, PI, SC, SF, CT, CTF, SD, DD, LF, FAD, RF, FL, WL, CL, CP, GM, P, RL, S, ST, SCN, SFD, TR, DR, FCS, ES, SEQ, WIN]
                
        #ComponentsDict = [ME, EE]

        #elementsMap = [architecture.MechanicalEquipment(), architecture.ElectricalEquipment(), architecture.CableTray()]



        for elements in ComponentsDict:
            if elements["Elements"]:
                elementsMap.append(elements['Elements'])
            else:
                continue

        for elements in ComponentsDict:
            if elements["Units"]:
                unitsMap.append(elements['Units'])
            else:
                continue

        #OUT = elementsMap

        t=Transaction(doc)

        


        for e in elementsMap:
            if len(e)!=0:
                for i in e:
                    #TEST.append(i)
                    
                    try:
                        # --- General Paramaters ---
                        elementId.append(i.Id)
                        elementCategory.append(i.Category.Name)

                        if (i.Category.Name == "Mechanical Equipment"):
                            unitVALUE.append(ME['Units'])
                        elif (i.Category.Name == "Electrical Equipment"):
                            unitVALUE.append(EE['Units'])
                        elif (i.Category.Name == "Duct Accessories"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Lighting Devices"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Cable Tray Fittings"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Air Terminals"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Data Devices"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Cable Trays"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Duct Fittings"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Duct Insulations"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Duct Fittings"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Fire Alarm Devices"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Lighting Devices"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Pipe Accessories"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Pipe Fittings"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Pipe Insulations"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Security Devices"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Parking"):
                            unitVALUE.append(DA['Units'])
                        elif (i.Category.Name == "Ducts"):
                            unitVALUE.append(D['Units'])                        
                        elif (i.Category.Name == "Structural Columns"):
                            unitVALUE.append(SC['Units'])                    
                        elif (i.Category.Name == "Structural Framing"):
                            unitVALUE.append(D['Units'])
                        elif (i.Category.Name == "Roofs"):
                            unitVALUE.append(RF['Units'])
                        
                        else:
                            unitVALUE.append('N/A')

                        elementGUID.append(i.UniqueId)			
                        famSymbol = sheets.GetElement(i.GetTypeId())
                        elementTYPE.append(Element.Name.__get__(famSymbol))			
                        elementFAM.append(str(famSymbol.FamilyName))
            
                        Manufacturer = ()
                        Model = ()
                        Description = ()
                        Name = ()	
                        
                        # --- Type Paramaters ---
                        
                        for p in famSymbol.GetOrderedParameters():
                            if p.Definition.Name == "Manufacturer":
                                Manufacturer = p
                            elif p.Definition.Name == "Model":
                                Model = p
                            elif p.Definition.Name == "Description":
                                Description = p
                            elif p.Definition.Name == "Name":
                                Name = p
                                
                        manufacturerVALUE.append(Manufacturer.AsString()) if Manufacturer else manufacturerVALUE.append(em1)
            
                        modelVALUE.append(Model.AsString()) if Model else modelVALUE.append(em1)
            
                        descriptionVALUE.append(str(Description.AsString()).encode('utf-8').strip()) if Description else descriptionVALUE.append(em1)
            
                        nameVALUE.append(Name.AsString()) if Name else nameVALUE.append(em1)
                        
                    except:
                        pass

                # --- Instance Paramaters ---
                
            if len(e)!= 0:		
                
                for i in e:
                
                    instanceParams = i.Parameters

                    MRTSYSTypeMark = ()	
                    MRTSYSDrgNo = ()
                    MRTSYSSystemNo = ()
                    MRTSYSItemNo = ()
                    MRTSAAssemblyRef = ()
                    MRTSAItemNo = ()
                    MRTSAProdRef = ()
                    MRTSASystemType = ()		
                    size = ()
                    assemblyName = ()
                    length = ()
                    WBSPAMMembermark = ()
                    
                    #instanceParameters = [MRTSYSTypeMark, MRTSYSDrgNo, MRTSYSSystemNo, MRTSYSItemNo, MRTSAAssemblyRef, MRTSAItemNo, MRTSAProdRef, MRTSASystemType, size, assemblyName]
                    
                    for p in instanceParams:
                        if p.Definition.Name == "MRT_SYS_Type Mark":
                            MRTSYSTypeMark = p
                        elif p.Definition.Name == "MRT_SYS_Drg No":
                            MRTSYSDrgNo = p
                        elif p.Definition.Name == "MRT_SYS_System No":
                            MRTSYSSystemNo = p
                        elif p.Definition.Name == "MRT_SYS_Item No":
                            MRTSYSItemNo = p
                        elif p.Definition.Name == "MRT_SA_Assembly Ref":
                            MRTSAAssemblyRef = p
                        elif p.Definition.Name == "MRT_SA_Item No":
                            MRTSAItemNo = p
                        elif p.Definition.Name == "MRT_SA_Prod Ref":
                            MRTSAProdRef = p
                        elif p.Definition.Name == "MRT_SA_System Type":
                            MRTSASystemType = p
                        elif p.Definition.Name == "Size":
                            size = p
                        elif p.Definition.Name == "Assembly Name":
                            assemblyName = p
                        elif p.Definition.Name == "Length":
                            length = p	
                        elif p.Definition.Name == "MRT_WBS PAM Membermark":
                            WBSPAMMembermark = p			

                    
                        
                    MRTSYSTypeMarkVALUE.append(MRTSYSTypeMark.AsString()) if MRTSYSTypeMark else MRTSYSTypeMarkVALUE.append(em1)
                        
                    MRTSYSDrgNoVALUE.append(MRTSYSDrgNo.AsString()) if MRTSYSDrgNo else MRTSYSDrgNoVALUE.append(em1)

                    MRTSAItemNoVALUE.append(MRTSYSSystemNo.AsString()) if MRTSYSSystemNo else MRTSYSystemNoVALUE.append(em1)

                    MRTSYSItemNoVALUE.append(MRTSYSItemNo.AsString()) if MRTSYSItemNo else MRTSYSItemNoVALUE.append(em1)

                    MRTSAAssemblyRefVALUE.append(MRTSAAssemblyRef.AsString()) if MRTSAAssemblyRef else MRTSAAssemblyRefVALUE.append(em1)	
                
                    MRTSAItemNoVALUE.append(MRTSAItemNo.AsString()) if MRTSAItemNo else MRTSAItemNoVALUE.append(em1)	

                    MRTSAProdRefVALUE.append(MRTSAProdRef.AsString()) if MRTSAProdRef else MRTSAProdRefVALUE.append(em1)

                    MRTSASystemTypeVALUE.append(MRTSASystemType.AsString()) if MRTSASystemType else MRTSASystemTypeVALUE.append(em1)

                    sizeVALUE.append(size.AsString()) if size else sizeVALUE.append(em1)

                    assemblyNameVALUE.append(assemblyName.AsString()) if assemblyName else assemblyNameVALUE.append(em1)

                    #(lengthVALUE.append(conv(length.AsDouble())), unitVALUE.append("m")) if length else (lengthVALUE.append(1), unitVALUE.append("EA"))

                    docTitle = sheets.Title    
                    modelName.append(docTitle)
                    
        TransferLog.append(['DONE'])
        DOCUMENTS.clear()
        elementsMap.clear()
        ComponentsDict.clear()
        ME.clear()
        EE.clear()
        D.clear()
        DA.clear()
        DF.clear()
        FD.clear()
        DI.clear()
        AT.clear()
        P.clear()
        PA.clear()
        PF.clear()
        PI.clear()
        SC.clear()
        SF.clear()
        CT.clear()
        CTF.clear()
        SD.clear()
        DD.clear()
        LF.clear()
        FAD.clear()
        RF.clear()
        FL.clear()
        WL.clear()
        CL.clear()
        CP.clear()
        GM.clear()
        P.clear()
        RL.clear()
        S.clear()
        ST.clear()
        SCN.clear()
        SFD.clear()
        TR.clear()
        DR.clear()
        FCS.clear()
        ES.clear()
        SEQ.clear()
        WIN.clear()        

        #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ EXPORT CSV ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

        #myLog   =  "M:\\10. User Documents\\Justin\\Export Test\\BOQ_Export_" + str(docTitle) + "_" + str(dateStamp) + "_" + str(timeStamp) + ".csv"      
    
    except:
        TransferLog.append(['Sheet Not found'])
        pass
    
    
docTitle = doc.Title

myLog   =  directory_raw + "\\BOQ_Export_" + str(docTitle) + "_" + str(dateStamp) + "_" + str(timeStamp) + ".csv"

header = "GUID", "Source Id", "Model Name", "Assembly Name", "Category Name", "Family Name", "Family Type", "Model", "Manufacturer", "Description", "Size", "Quantity", "Unit", "MRT_SA_Assembly Ref", "MRT_SA_Item No", "MRT_SA_Prod Ref", "MRT_SA_System Type", "MRT_SYS_Type Mark", "MRT_SYS_Drg NO", "MRT_SYS_System No", "MRT_SYS_Item No" 

dataZip = izip_longest(elementGUID, elementId, modelName, assemblyNameVALUE, elementCategory, elementFAM, elementTYPE, modelVALUE, manufacturerVALUE, descriptionVALUE, sizeVALUE,lengthVALUE, unitVALUE, MRTSAAssemblyRefVALUE, MRTSAItemNoVALUE, MRTSAProdRefVALUE, MRTSASystemTypeVALUE, MRTSYSTypeMarkVALUE, MRTSYSDrgNoVALUE, MRTSYSystemNoVALUE, MRTSYSItemNoVALUE)

data = [i for i in dataZip]


with open(myLog, 'wb') as f:
    #using csv.writer method from CSV package
    write = csv.writer(f, dialect = csv.excel)
    write.writerow(header)
    for row in data:
        write.writerow(row)



OUT = TransferLog
        #OUT = mArc
        # Assign your output to the OUT variable.

#TransactionManager.Instance.TransactionTaskDone()









