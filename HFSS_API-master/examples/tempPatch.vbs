Dim oHfssApp
Dim oDesktop
Dim oProject
Dim oDesign
Dim oEditor
Dim oModule

Set oHfssApp  = CreateObject("AnsoftHfss.HfssScriptInterface")
Set oDesktop = oHfssApp.GetAppDesktop()
oDesktop.RestoreWindow
oDesktop.NewProject
Set oProject = oDesktop.GetActiveProject

oProject.InsertDesign "HFSS", "AntenaParcheMama1", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("AntenaParcheMama1")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "-0.016500meter", _
"YStart:=", "-0.017500meter", _
"ZStart:=", "0.000000meter", _
"Width:=", "0.033000meter", _
"Height:=", "0.035000meter", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "Patch", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "-0.014833meter", _
"YStart:=", "-0.017500meter", _
"ZStart:=", "0.000000meter", _
"Width:=", "0.005000meter", _
"Height:=", "-0.002500meter", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "Acople", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "-0.012618meter", _
"YStart:=", "-0.020000meter", _
"ZStart:=", "0.000000meter", _
"Width:=", "0.000570meter", _
"Height:=", "-0.015000meter", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "Feed", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.Unite  _
Array("NAME:Selections", _
"Selections:=", "Patch,Acople,Feed"), _
Array("NAME:UniteParameters", "KeepOriginals:=", false)

Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignPerfectE _
Array("NAME:PatchMetal", _
"InfGroundPlane:=", false, _
"Objects:=", _
Array("Patch"))

oEditor.ChangeProperty _
	Array("NAME:AllTabs", _
		Array("NAME:Geometry3DAttributeTab", _
			Array("NAME:PropServers", "Patch"), _
			Array("NAME:ChangedProps",  _
				Array("NAME:Color", "R:=", 255, "G:=", 255, "B:=", 0) _
			) _
		) _
	) 

oEditor.ChangeProperty _
Array("NAME:AllTabs", _
	Array("NAME:Geometry3DAttributeTab", _
		Array("NAME:PropServers","Patch"), _
		Array("NAME:ChangedProps", _
			Array("NAME:Transparent", "Value:=",  0.000000) _
			) _
		) _
	)

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "-0.012618meter", _
"YStart:=", "-0.035000meter", _
"ZStart:=", "0.000000meter", _
"Width:=", "-0.000250meter", _
"Height:=", "0.000570meter", _
"WhichAxis:=", "Y"), _
Array("NAME:Attributes", _
"Name:=", "Port", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignLumpedPort _
Array("NAME:LPort", _
      Array("NAME:Modes", _
             Array("NAME:Mode1", _
                   "ModeNum:=", 1, _
                   "UseIntLine:=", true, _
                   Array("NAME:IntLine", _
                          "Start:=", Array("-0.012333meter", "-0.035000meter", "0.000000meter"), _
                          "End:=",   Array("-0.012333meter", "-0.035000meter", "-0.000250meter") _
                         ), _
                   "CharImp:=", "Zpi" _
                   ) _
            ), _
      "Resistance:=", "50.000000Ohm", _
      "Reactance:=", "0.000000Ohm", _
      "Objects:=", Array("Port") _
      )

oEditor.CreateBox _
Array("NAME:BoxParameters", _
"XPosition:=", "-0.066000meter", _
"YPosition:=", "-0.035000meter", _
"ZPosition:=", "0.000000meter", _
"XSize:=", "0.132000meter", _
"YSize:=", "0.070000meter", _
"ZSize:=", "-0.000250meter"), _
Array("NAME:Attributes", _
"Name:=", "Substrate", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 0.75, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.AssignMaterial _
	Array("NAME:Selections", _
		"Selections:=", "Substrate"), _
	Array("NAME:Attributes", _
		"MaterialName:=", "Dupont Type 100 HN Film (tm)", _
		"SolveInside:=", true)

oEditor.ChangeProperty _
	Array("NAME:AllTabs", _
		Array("NAME:Geometry3DAttributeTab", _
			Array("NAME:PropServers", "Substrate"), _
			Array("NAME:ChangedProps",  _
				Array("NAME:Color", "R:=", 0, "G:=", 128, "B:=", 0) _
			) _
		) _
	) 

oEditor.ChangeProperty _
Array("NAME:AllTabs", _
	Array("NAME:Geometry3DAttributeTab", _
		Array("NAME:PropServers","Substrate"), _
		Array("NAME:ChangedProps", _
			Array("NAME:Transparent", "Value:=",  0.200000) _
			) _
		) _
	)

oEditor.CreateBox _
Array("NAME:BoxParameters", _
"XPosition:=", "-0.183673meter", _
"YPosition:=", "-0.091837meter", _
"ZPosition:=", "-0.000500meter", _
"XSize:=", "0.367347meter", _
"YSize:=", "0.367347meter", _
"ZSize:=", "0.061724meter"), _
Array("NAME:Attributes", _
"Name:=", "AirBox", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 0.75, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignRadiation _
Array("NAME:Radiation", _
"Objects:=", Array("AirBox"))

oEditor.ChangeProperty _
Array("NAME:AllTabs", _
	Array("NAME:Geometry3DAttributeTab", _
		Array("NAME:PropServers","AirBox"), _
		Array("NAME:ChangedProps", _
			Array("NAME:Transparent", "Value:=",  0.950000) _
			) _
		) _
	)

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "-0.066000meter", _
"YStart:=", "-0.035000meter", _
"ZStart:=", "-0.000250meter", _
"Width:=", "0.132000meter", _
"Height:=", "0.070000meter", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "GroundPlane", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignPerfectE _
Array("NAME:GND", _
"InfGroundPlane:=", false, _
"Objects:=", _
Array("GroundPlane"))

oEditor.ChangeProperty _
	Array("NAME:AllTabs", _
		Array("NAME:Geometry3DAttributeTab", _
			Array("NAME:PropServers", "GroundPlane"), _
			Array("NAME:ChangedProps",  _
				Array("NAME:Color", "R:=", 30, "G:=", 128, "B:=", 30) _
			) _
		) _
	) 

oEditor.ChangeProperty _
Array("NAME:AllTabs", _
	Array("NAME:Geometry3DAttributeTab", _
		Array("NAME:PropServers","GroundPlane"), _
		Array("NAME:ChangedProps", _
			Array("NAME:Transparent", "Value:=",  0.200000) _
			) _
		) _
	)

oEditor.Move _
Array("NAME:Selections", _
"Selections:=", "Patch,Port"), _
Array("NAME:TranslateParameters", _
"TranslateVectorX:=", "0.030612meter", _
"TranslateVectorY:=", "0.000000meter", _
"TranslateVectorZ:=", "0.000000meter")

oEditor.DuplicateAlongLine _
Array("NAME:Selections", _
"Selections:=", "Patch,Port"), _
Array("NAME:DuplicateToAlongLineParameters", _
"XComponent:=", "-0.061224meter", _
"YComponent:=", "0.000000meter", _
"ZComponent:=", "0.000000meter", _
"NumClones:=", 2), _
Array("NAME:Options", _
"DuplicateBoundaries:=", true)
oEditor.DuplicateAroundAxis _
	Array("NAME:Selections", "Selections:=", "Patch,Substrate,Port,Patch_1,Port_1,GroundPlane","NewPartsModelFlag:=", "Model"), _
	Array("NAME:DuplicateAroundAxisParameters", "CreateNewObjects:=", true, "WhichAxis:=", _
"Z", "AngleStr:=",_
 "90.000000deg", "NumClones:=", "2"), _
Array("NAME:Options", "DuplicateAssignments:=",true), Array("CreateGroupsForNewObjects:=", false) 

oEditor.DuplicateAroundAxis _
	Array("NAME:Selections", "Selections:=", "Patch,Substrate,Port,Patch_1,Port_1,GroundPlane","NewPartsModelFlag:=", "Model"), _
	Array("NAME:DuplicateAroundAxisParameters", "CreateNewObjects:=", true, "WhichAxis:=", _
"Z", "AngleStr:=",_
 "180.000000deg", "NumClones:=", "2"), _
Array("NAME:Options", "DuplicateAssignments:=",true), Array("CreateGroupsForNewObjects:=", false) 

oEditor.DuplicateAroundAxis _
	Array("NAME:Selections", "Selections:=", "Patch,Substrate,Port,Patch_1,Port_1,GroundPlane","NewPartsModelFlag:=", "Model"), _
	Array("NAME:DuplicateAroundAxisParameters", "CreateNewObjects:=", true, "WhichAxis:=", _
"Z", "AngleStr:=",_
 "270.000000deg", "NumClones:=", "2"), _
Array("NAME:Options", "DuplicateAssignments:=",true), Array("CreateGroupsForNewObjects:=", false) 


oEditor.Move _
Array("NAME:Selections", _
"Selections:=", "Substrate_1,Patch_2,Patch_1_1,Port_2,Port_1_1,GroundPlane_1"), _
Array("NAME:TranslateParameters", _
"TranslateVectorX:=", "-0.105000meter", _
"TranslateVectorY:=", "0.106000meter", _
"TranslateVectorZ:=", "0.000000meter")

oEditor.Move _
Array("NAME:Selections", _
"Selections:=", "Substrate_1,Patch_2,Patch_1_1,Port_2,Port_1_1,GroundPlane_1"), _
Array("NAME:TranslateParameters", _
"TranslateVectorX:=", "0.212000meter", _
"TranslateVectorY:=", "0.000000meter", _
"TranslateVectorZ:=", "0.000000meter")

oEditor.Move _
Array("NAME:Selections", _
"Selections:=", "Substrate_2,Patch_3,Patch_1_2,Port_3,Port_1_2,GroundPlane_2"), _
Array("NAME:TranslateParameters", _
"TranslateVectorX:=", "0.000000meter", _
"TranslateVectorY:=", "0.212000meter", _
"TranslateVectorZ:=", "0.000000meter")

oEditor.Move _
Array("NAME:Selections", _
"Selections:=", "Substrate_3,Patch_4,Patch_1_3,Port_4,Port_1_3,GroundPlane_3"), _
Array("NAME:TranslateParameters", _
"TranslateVectorX:=", "0.105000meter", _
"TranslateVectorY:=", "0.106000meter", _
"TranslateVectorZ:=", "0.000000meter")

oEditor.Move _
Array("NAME:Selections", _
"Selections:=", "Substrate_3,Patch_4,Patch_1_3,Port_4,Port_1_3,GroundPlane_3"), _
Array("NAME:TranslateParameters", _
"TranslateVectorX:=", "-0.212000meter", _
"TranslateVectorY:=", "0.000000meter", _
"TranslateVectorZ:=", "0.000000meter")

Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "HfssDriven", _
Array("NAME:Setup2_45GHz", _
"Frequency:=", "2.450000GHz", _
"PortsOnly:=", false, _
"maxDeltaS:=", 0.020000, _
"UseMatrixConv:=", false, _
"MaximumPasses:=", 25, _
"MinimumPasses:=", 1, _
"MinimumConvergedPasses:=", 1, _
"PercentRefinement:=", 20, _
"ReducedSolutionBasis:=", false, _
"DoLambdaRefine:=", true, _
"DoMaterialLambda:=", true, _
"Target:=", 0.3333, _
"PortAccuracy:=", 2, _
"SetPortMinMaxTri:=", false)

Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertDrivenSweep _
"Setup2_45GHz", _
Array("NAME:Sweep_1a5GHz", _
"Type:=", "Interpolating", _
"InterpTolerance:=", 0.500000, _
"InterpMaxSolns:=", 101, _
"SetupType:=", "LinearCount", _
"StartFreq:=", "1.000000GHz", _
"StopFreq:=", "5.000000GHz", _
"Count:=", 201, _
"SaveFields:=", false, _
"ExtrapToDC:=", false)

oProject.SaveAs _
    "C:\Users\Asus\Documents\Ansoft\antenaDMamaArray1.aedt", _
    true
