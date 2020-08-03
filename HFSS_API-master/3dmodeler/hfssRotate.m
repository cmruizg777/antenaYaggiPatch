% ----------------------------------------------------------------------------
% function hfssMove(fid, ObjList, tVector, Units)
% 
% Description :
% -------------
% Creates the VB Script necessary to move a given set of objects.
%
% Parameters :
% ------------
% fid     - file identifier of the HFSS script file.
% ObjList - a cell array of objects that need to be moved.
% tVector - a translation vector that defines the motion (specify as 
%           [tx, ty, tz]).
% Units   - units for the transalation vector (use 'in', 'mm', 'meter' or 
%           anything else defined in HFSS).
% 
% Note :
% ------
%
% Example :
% ---------
% fid = fopen('myantenna.vbs', 'wt');
% ... create some objects here ...
% hfssMove(fid, {'Dipole1', 'Dipole2'}, [0, d, 0], 'in');
% ----------------------------------------------------------------------------

% ----------------------------------------------------------------------------
% This file is part of HFSS-MATLAB-API.
%
% HFSS-MATLAB-API is free software; you can redistribute it and/or modify it 
% under the terms of the GNU General Public License as published by the Free 
% Software Foundation; either version 2 of the License, or (at your option) 
% any later version.
%
% HFSS-MATLAB-API is distributed in the hope that it will be useful, but 
% WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY 
% or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License 
% for more details.
%
% You should have received a copy of the GNU General Public License along with
% Foobar; if not, write to the Free Software Foundation, Inc., 59 Temple 
% Place, Suite 330, Boston, MA  02111-1307  USA
%
% Copyright 2004, Vijay Ramasami (rvc@ku.edu)
% ----------------------------------------------------------------------------
function hfssRotate(fid, ObjectList, angle, axis)

nObjects = length(ObjectList);

% Preamble.
fprintf(fid, '\n');
fprintf(fid, 'oEditor.Rotate _\n');
fprintf(fid, 'Array("NAME:Selections", _\n');

% Object Selections.
fprintf(fid, '"Selections:=", "');
for iObj = 1:nObjects,
	fprintf(fid, '%s', ObjectList{iObj});
	if (iObj ~= nObjects)
		fprintf(fid, ',');
	end;
end;
fprintf(fid, '", _\n');
fprintf(fid, '"NewPartsModelFlag:=",  ');
fprintf(fid, '"Model"), Array("NAME:RotateParameters", "RotateAxis:=", "%s", "RotateAngle:=",  ', axis);
fprintf(fid, '"%fdeg")', angle);

