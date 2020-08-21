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

oProject.InsertDesign "HFSS", "yaggi", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("yaggi")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "0.000000mm", _
"YStart:=", "-0.500000mm", _
"ZStart:=", "0.000000mm", _
"Width:=", "2.500000mm", _
"Height:=", "1.000000mm", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "r1", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "2.500000mm", _
"YStart:=", "-0.270000mm", _
"ZStart:=", "0.000000mm", _
"Width:=", "1.910000mm", _
"Height:=", "0.540000mm", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "r2", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "4.410000mm", _
"YStart:=", "-0.325000mm", _
"ZStart:=", "0.000000mm", _
"Width:=", "1.640000mm", _
"Height:=", "0.650000mm", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "r3", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "6.050000mm", _
"YStart:=", "-2.595000mm", _
"ZStart:=", "0.000000mm", _
"Width:=", "3.000000mm", _
"Height:=", "5.190000mm", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "r4", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "9.050000mm", _
"YStart:=", "-0.650000mm", _
"ZStart:=", "0.000000mm", _
"Width:=", "1.000000mm", _
"Height:=", "1.300000mm", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "r5", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "10.050000mm", _
"YStart:=", "-2.595000mm", _
"ZStart:=", "0.000000mm", _
"Width:=", "0.800000mm", _
"Height:=", "5.190000mm", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "r6", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "13.850000mm", _
"YStart:=", "-1.375000mm", _
"ZStart:=", "0.000000mm", _
"Width:=", "0.720000mm", _
"Height:=", "2.750000mm", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "r7", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "17.650000mm", _
"YStart:=", "-1.375000mm", _
"ZStart:=", "0.000000mm", _
"Width:=", "0.720000mm", _
"Height:=", "2.750000mm", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "r8", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "21.450000mm", _
"YStart:=", "-1.375000mm", _
"ZStart:=", "0.000000mm", _
"Width:=", "0.720000mm", _
"Height:=", "2.750000mm", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "r9", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.CreateBox _
Array("NAME:BoxParameters", _
"XPosition:=", "0.000000mm", _
"YPosition:=", "-10.000000mm", _
"ZPosition:=", "0.000000mm", _
"XSize:=", "22.550000mm", _
"YSize:=", "20.000000mm", _
"ZSize:=", "-0.500000mm"), _
Array("NAME:Attributes", _
"Name:=", "Sustrato", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 0.75, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.AssignMaterial _
	Array("NAME:Selections", _
		"Selections:=", "Sustrato"), _
	Array("NAME:Attributes", _
		"MaterialName:=", "FR4_epoxy", _
		"SolveInside:=", true)

oEditor.Unite  _
Array("NAME:Selections", _
"Selections:=", "r1,r2,r3,r4,r5,r6,r7,r8,r9"), _
Array("NAME:UniteParameters", "KeepOriginals:=", false)

oEditor.ChangeProperty _
	Array("NAME:AllTabs", _
		Array("NAME:Geometry3DAttributeTab", _
			Array("NAME:PropServers", "r1"), _
			Array("NAME:ChangedProps",  _
				Array("NAME:Color", "R:=", 220, "G:=", 220, "B:=", 10) _
			) _
		) _
	) 

oEditor.ChangeProperty _
Array("NAME:AllTabs", _
	Array("NAME:Geometry3DAttributeTab", _
		Array("NAME:PropServers","r1"), _
		Array("NAME:ChangedProps", _
			Array("NAME:Transparent", "Value:=",  0.000000) _
			) _
		) _
	)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignFiniteCond Array("NAME:Patch", "Objects:=", Array( _
  "r1"), "UseMaterial:=", false, "Roughness:=",  _
  "0um", "UseThickness:=", false, "Thickness:=", false, "InfGroundPlane:=",  _
  false)


oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "7.550000mm", _
"YStart:=", "-0.150000mm", _
"ZStart:=", "0.000000mm", _
"Width:=", "3.300000mm", _
"Height:=", "0.300000mm", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "e1", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.ChangeProperty _
	Array("NAME:AllTabs", _
		Array("NAME:Geometry3DAttributeTab", _
			Array("NAME:PropServers", "e1"), _
			Array("NAME:ChangedProps",  _
				Array("NAME:Color", "R:=", 220, "G:=", 0, "B:=", 0) _
			) _
		) _
	) 

oEditor.CreatePolyline _
	Array("NAME:PolylineParameters", "IsPolylineCovered:=", true, "IsPolylineClosed:=", true, _
		Array("NAME:PolylinePoints", _
			Array("NAME:PLPoint", "X:=", "9.0500mm", "Y:=", "0.0000mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "7.5500mm", "Y:=", "-1.5000mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "6.0500mm", "Y:=", "0.0000mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "7.5500mm", "Y:=", "1.5000mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "9.0500mm",   "Y:=", "0.0000mm",   "Z:=", "0.0000mm")_
			), _ 
		Array("NAME:PolylineSegments", _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2), _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 1, "NoOfPoints:=", 2), _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 2, "NoOfPoints:=", 2), _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 3, "NoOfPoints:=", 2) _
			) _
		), _
Array("NAME:Attributes", _
"Name:=", "e2", _
"Flags:=", "", _
"Color:=", "(0 0 0)", _
"Transparency:=", 0.800000, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.ChangeProperty _
	Array("NAME:AllTabs", _
		Array("NAME:Geometry3DAttributeTab", _
			Array("NAME:PropServers", "e2"), _
			Array("NAME:ChangedProps",  _
				Array("NAME:Color", "R:=", 220, "G:=", 0, "B:=", 0) _
			) _
		) _
	) 

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "7.050000mm", _
"YStart:=", "-1.595000mm", _
"ZStart:=", "0.000000mm", _
"Width:=", "1.000000mm", _
"Height:=", "3.190000mm", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "e3", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.ChangeProperty _
	Array("NAME:AllTabs", _
		Array("NAME:Geometry3DAttributeTab", _
			Array("NAME:PropServers", "e3"), _
			Array("NAME:ChangedProps",  _
				Array("NAME:Color", "R:=", 220, "G:=", 0, "B:=", 0) _
			) _
		) _
	) 

oEditor.CreatePolyline _
	Array("NAME:PolylineParameters", "IsPolylineCovered:=", true, "IsPolylineClosed:=", true, _
		Array("NAME:PolylinePoints", _
			Array("NAME:PLPoint", "X:=", "9.0500mm", "Y:=", "-2.5950mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "9.0500mm", "Y:=", "-1.5950mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "8.0500mm", "Y:=", "-2.5950mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "9.0500mm",   "Y:=", "-2.5950mm",   "Z:=", "0.0000mm")_
			), _ 
		Array("NAME:PolylineSegments", _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2), _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 1, "NoOfPoints:=", 2), _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 2, "NoOfPoints:=", 2) _
			) _
		), _
Array("NAME:Attributes", _
"Name:=", "e4", _
"Flags:=", "", _
"Color:=", "(0 0 0)", _
"Transparency:=", 0.800000, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.CreatePolyline _
	Array("NAME:PolylineParameters", "IsPolylineCovered:=", true, "IsPolylineClosed:=", true, _
		Array("NAME:PolylinePoints", _
			Array("NAME:PLPoint", "X:=", "9.0500mm", "Y:=", "2.5950mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "9.0500mm", "Y:=", "1.5950mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "8.0500mm", "Y:=", "2.5950mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "9.0500mm",   "Y:=", "2.5950mm",   "Z:=", "0.0000mm")_
			), _ 
		Array("NAME:PolylineSegments", _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2), _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 1, "NoOfPoints:=", 2), _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 2, "NoOfPoints:=", 2) _
			) _
		), _
Array("NAME:Attributes", _
"Name:=", "e5", _
"Flags:=", "", _
"Color:=", "(0 0 0)", _
"Transparency:=", 0.800000, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.CreatePolyline _
	Array("NAME:PolylineParameters", "IsPolylineCovered:=", true, "IsPolylineClosed:=", true, _
		Array("NAME:PolylinePoints", _
			Array("NAME:PLPoint", "X:=", "6.0500mm", "Y:=", "-2.5950mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "7.0500mm", "Y:=", "-2.5950mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "6.0500mm", "Y:=", "-1.5950mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "6.0500mm",   "Y:=", "-2.5950mm",   "Z:=", "0.0000mm")_
			), _ 
		Array("NAME:PolylineSegments", _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2), _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 1, "NoOfPoints:=", 2), _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 2, "NoOfPoints:=", 2) _
			) _
		), _
Array("NAME:Attributes", _
"Name:=", "e6", _
"Flags:=", "", _
"Color:=", "(0 0 0)", _
"Transparency:=", 0.800000, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.CreatePolyline _
	Array("NAME:PolylineParameters", "IsPolylineCovered:=", true, "IsPolylineClosed:=", true, _
		Array("NAME:PolylinePoints", _
			Array("NAME:PLPoint", "X:=", "6.0500mm", "Y:=", "2.5950mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "7.0500mm", "Y:=", "2.5950mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "6.0500mm", "Y:=", "1.5950mm", "Z:=", "0.0000mm"), _
			Array("NAME:PLPoint", "X:=", "6.0500mm",   "Y:=", "2.5950mm",   "Z:=", "0.0000mm")_
			), _ 
		Array("NAME:PolylineSegments", _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 0, "NoOfPoints:=", 2), _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 1, "NoOfPoints:=", 2), _
			Array("NAME:PLSegment", "SegmentType:=",  "Line", "StartIndex:=", 2, "NoOfPoints:=", 2) _
			) _
		), _
Array("NAME:Attributes", _
"Name:=", "e7", _
"Flags:=", "", _
"Color:=", "(0 0 0)", _
"Transparency:=", 0.800000, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.Subtract _
Array("NAME:Selections", _
"Blank Parts:=", _
"r1", _
"Tool Parts:=", _
"e1,e2,e3,e4,e5,e6,e7"), _
Array("NAME:SubtractParameters", _
"KeepOriginals:=", false) 

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "0.000000mm", _
"YStart:=", "-10.000000mm", _
"ZStart:=", "-0.500000mm", _
"Width:=", "9.050000mm", _
"Height:=", "20.000000mm", _
"WhichAxis:=", "Z"), _
Array("NAME:Attributes", _
"Name:=", "gnd", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.ChangeProperty _
	Array("NAME:AllTabs", _
		Array("NAME:Geometry3DAttributeTab", _
			Array("NAME:PropServers", "gnd"), _
			Array("NAME:ChangedProps",  _
				Array("NAME:Color", "R:=", 10, "G:=", 220, "B:=", 10) _
			) _
		) _
	) 

oEditor.ChangeProperty _
Array("NAME:AllTabs", _
	Array("NAME:Geometry3DAttributeTab", _
		Array("NAME:PropServers","gnd"), _
		Array("NAME:ChangedProps", _
			Array("NAME:Transparent", "Value:=",  0.000000) _
			) _
		) _
	)

Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignPerfectE _
Array("NAME:GND", _
"InfGroundPlane:=", true, _
"Objects:=", _
Array("gnd"))

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "0.000000mm", _
"YStart:=", "-0.500000mm", _
"ZStart:=", "-0.500000mm", _
"Width:=", "1.000000mm", _
"Height:=", "0.500000mm", _
"WhichAxis:=", "X"), _
Array("NAME:Attributes", _
"Name:=", "port", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 5.000000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.ChangeProperty _
	Array("NAME:AllTabs", _
		Array("NAME:Geometry3DAttributeTab", _
			Array("NAME:PropServers", "port"), _
			Array("NAME:ChangedProps",  _
				Array("NAME:Color", "R:=", 235, "G:=", 10, "B:=", 10) _
			) _
		) _
	) 

oEditor.ChangeProperty _
Array("NAME:AllTabs", _
	Array("NAME:Geometry3DAttributeTab", _
		Array("NAME:PropServers","port"), _
		Array("NAME:ChangedProps", _
			Array("NAME:Transparent", "Value:=",  0.000000) _
			) _
		) _
	)

Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignLumpedPort _
Array("NAME:puerto", _
      Array("NAME:Modes", _
             Array("NAME:Mode1", _
                   "ModeNum:=", 1, _
                   "UseIntLine:=", true, _
                   Array("NAME:IntLine", _
                          "Start:=", Array("0.000000mm", "0.000000mm", "-0.250000mm"), _
                          "End:=",   Array("0.000000mm", "0.000000mm", "0.000000mm") _
                         ), _
                   "CharImp:=", "Zpi" _
                   ) _
            ), _
      "Resistance:=", "50.000000Ohm", _
      "Reactance:=", "0.000000Ohm", _
      "Objects:=", Array("port") _
      )

oEditor.CreateBox _
Array("NAME:BoxParameters", _
"XPosition:=", "-1.000000mm", _
"YPosition:=", "-11.000000mm", _
"ZPosition:=", "-1.000000mm", _
"XSize:=", "24.550000mm", _
"YSize:=", "22.000000mm", _
"ZSize:=", "10.000000mm"), _
Array("NAME:Attributes", _
"Name:=", "box2", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 0.75, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.ChangeProperty _
Array("NAME:AllTabs", _
	Array("NAME:Geometry3DAttributeTab", _
		Array("NAME:PropServers","box2"), _
		Array("NAME:ChangedProps", _
			Array("NAME:Transparent", "Value:=",  0.950000) _
			) _
		) _
	)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignRadiation _
Array("NAME:rad1", _
"Objects:=", Array("box2"))

Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "HfssDriven", _
Array("NAME:yaggi_uda", _
"Frequency:=", "5.000000GHz", _
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
"yaggi_uda", _
Array("NAME:SWEEP1", _
"Type:=", "Interpolating", _
"InterpTolerance:=", 0.500000, _
"InterpMaxSolns:=", 101, _
"SetupType:=", "LinearCount", _
"StartFreq:=", "3.000000GHz", _
"StopFreq:=", "7.000000GHz", _
"Count:=", 101, _
"SaveFields:=", false, _
"ExtrapToDC:=", false)

Set oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport "Report 1", _
"Modal S Parameter", _
"Rectangular Plot", _
"yaggi_uda : Sweep1", _
Array("Domain:=", "Sweep"), _
Array("Freq:=", Array("All")), _
Array("X Component:=", "Freq", _
"Y Component:=", Array("db(S(puerto,puerto))")), _
Array()

oProject.SaveAs _
    "D:\MATLAB\Oscar\Yaggi\yaggi.vbs.aedt", _
    true
