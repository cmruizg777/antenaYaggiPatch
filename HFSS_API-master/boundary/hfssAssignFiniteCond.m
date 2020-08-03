%function hfssAssignFiniteCond(fid, Name, BoxObject, material, thickness, igp)
function hfssAssignFiniteCond(fid, Name, BoxObject, thickness, igp)
fprintf(fid, 'Set oModule = oDesign.GetModule("BoundarySetup")\n');
fprintf(fid, 'oModule.AssignFiniteCond Array("NAME:%s", "Objects:=", Array( _\n',Name);
%fprintf(fid, '  "%s"), "UseMaterial:=", true, "Material:=", "%s", "Roughness:=",  _\n', BoxObject, material);
fprintf(fid, '  "%s"), "UseMaterial:=", false, "Roughness:=",  _\n', BoxObject);
fprintf(fid, '  "0um", "UseThickness:=", false, "Thickness:=", %s, "InfGroundPlane:=",  _\n', thickness);
if(igp)
    fprintf(fid, '  true)\n\n'); 
else
    fprintf(fid, '  false)\n\n'); 
end