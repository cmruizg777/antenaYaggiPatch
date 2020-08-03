function hfssRegularPolygon(fid, name, center, start, n, plane )

% arguments processing.
if (nargin < 3)
	error('non-enough arguments !');
end;


fprintf(fid, '\n');
fprintf(fid, 'oEditor.CreateRegularPolygon Array("NAME:RegularPolygonParameters", "IsCovered:=",  _ \n');
fprintf(fid, 'true, "XCenter:=", "%fmm", "YCenter:=", "%fmm", "ZCenter:=", "%fmm", "XStart:=",  _ \n',center(1),center(2),center(3));
fprintf(fid, '"%fmm", "YStart:=", "%fmm", "ZStart:=", "%fmm", "NumSides:=", "%f", "WhichAxis:=",  _ \n',start(1),start(2),start(3),n);
fprintf(fid, '"%s"), Array("NAME:Attributes", "Name:=", "%s", "Flags:=", "", "Color:=",  _ \n',plane,name);
fprintf(fid, '"(143 175 143)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=",  _ \n');
fprintf(fid, '"", "MaterialValue:=", "" & Chr(34) & "vacuum" & Chr(34) & "", "SurfaceMaterialValue:=",  _ \n');
fprintf(fid, '"" & Chr(34) & "" & Chr(34) & "", "SolveInside:=", true, "IsMaterialEditable:=",  _ \n');
fprintf(fid, 'true, "UseMaterialAppearance:=", false, "IsLightweight:=", false) \n');


