% ----------------------------------------------------------------------------
% This file is part of HFSS-MATLAB-API.
%
% HFSS-MATLAB-API is free software; you can redistribute it and/or modify it 
% under the terms of the GNU General Public License as published by the Free 
% Software Foundation; either version 2 of the License, or (at your option) 
% any later version.
%

% ----------------------------------------------------------------------------
function hfssCone(fid, Name, Axis, Center, BRadius, TRadius, Height, Units)
	% Creates the VB script necessary to model a cone in HFSS.
	%
	% Parameters :
	% fid:		file identifier of the HFSS script file.
	% Name:		name of the cylinder (in HFSS).
	% Center:	center of the cylinder (specify as [x, y, z]). This is also the 
	%           starting point of the cone.
	% Axis:		axis of the cylinder (specify as 'X', 'Y', or 'Z').
	% BRadius:	bottom radius of the cone (scalar).
	% TRadius:	top radius of the cone (scalar).
	% Height:	height of the cylidner (from the point specified by Center).
	% Units:	specify as 'in', 'mm', 'meter' or anything else defined in HFSS.
	%
	% Example :
	% @code
	% fid = fopen('myantenna.vbs', 'wt');
	% ... 
	% hfssCone(fid, 'Cone1', 'Z', [0, 0, 0], 0.1, 10, 'mm');\
	% @endcode

	% Cone Parameters.
	fprintf(fid, '\n');
	fprintf(fid, 'oEditor.CreateCone _\n');
	fprintf(fid, 'Array("NAME:ConeParameters", _\n');
	fprintf(fid, '"XCenter:=", "%f%s", _\n', Center(1), Units);
	fprintf(fid, '"YCenter:=", "%f%s", _\n', Center(2), Units);
	fprintf(fid, '"ZCenter:=", "%f%s", _\n', Center(3), Units);
	fprintf(fid, '"WhichAxis:=", "%s", _\n', upper(Axis));
	fprintf(fid, '"Height:=", "%f%s", _\n', Height, Units);
	fprintf(fid, '"BottomRadius:=", "%f%s", _\n', BRadius, Units);
	fprintf(fid, '"TopRadius:=", "%f%s"), ', TRadius, Units);
		
	

	% Cone Properties.
	fprintf(fid, 'Array("NAME:Attributes", _\n'); 
	fprintf(fid, '"Name:=", "%s", _\n', Name);
	fprintf(fid, '"Flags:=", "", _\n');
	fprintf(fid, '"Color:=", "(132 132 193)", _\n');
	fprintf(fid, '"Transparency:=", 0, _\n');
	fprintf(fid, '"PartCoordinateSystem:=", "Global", _\n');
	fprintf(fid, '"MaterialName:=", "vacuum", _\n');
	fprintf(fid, '"SolveInside:=", true)\n');
	fprintf(fid, '\n');
	