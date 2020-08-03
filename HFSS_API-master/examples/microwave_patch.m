%*******************************************************************
% Script para generar una antena array MxN.
%*******************************************************************

% Limpiar el workspace y la ventana de comandos.
clear all;
close all;

% Agregar al path las funciones del API.
addpath('../boundary/');
addpath('../3dmodeler/');
addpath('../analysis/');
addpath('../general/');

% Ruta de ejecuci�n de HFSS. (COLOCAR SU PROPIA RUTA.
hfssExePath = 'D:\Programas\ANSYS\AnsysEM19.2\Win64\ansysedt.exe';

% C�lculo de lambda a partir de la frecuencia de operaci�n.
c = 3e8;
fC = 3.75e9;
Lambda = c/fC;

% Dimensiones del parche.
Wy = 24*1e-3;
Lx = 24*1e-3;

% Par�metros del array.
Nx = 3;
Ny = 3;
dAnt = Lambda/2;

% Dimensiones de la alimentaci�n uStrip.
Wms = (200/1000)*0.0254;
Lfeed = Lx/6;

% Par�metros del substrato.
Hsub = (60/1000)*0.0254;
WsubX = Nx*dAnt + Lambda/2;
WsubY = Ny*dAnt + Lambda/2;

% Par�metros del puerto.
PortH = Hsub;
PortW = Wms;

% Par�metros caja de radiaci�n.
AirX = Nx*dAnt + Lambda/2;
AirY = Ny*dAnt + Lambda/2;
AirZ = Lambda/4 + Hsub;
	
% Par�metros de soluci�n (Solution Setup) (GHz).
fSolve = fC/1e9;
fStart = 3;
fStop  = 4.5;

% Abre un script temporal y almacena las �rdenes.
tmpScriptFile = 'tempPatch.vbs';
fid = fopen(tmpScriptFile, 'wt');

% Crea un nuevo proyecto en HFSS.
hfssNewProject(fid);
hfssInsertDesign(fid, 'NxN_uStrip_Patch'); % Nombre.

% Genera el parche y la alimentaci�n.
hfssRectangle(fid, 'Patch', 'Z', [-Lx/2, -Wy/2, 0], Lx, Wy, 'meter');
hfssRectangle(fid, 'Feed', 'Z', [-Lx/2, -Wms/2, 0], -Lfeed, Wms, 'meter');
% Une el parche y la alimentaci�n
hfssUnite(fid, {'Patch', 'Feed'});
% Asigna l�mite al parche (Perfect E).
hfssAssignPE(fid, 'PatchMetal', {'Patch'});
hfssSetColor(fid, 'Patch', [128, 128, 0]); % Color.
hfssSetTransparency(fid, {'Patch'}, 0); % Transparencia.

% Genera un lumped port y lo asigna.
hfssRectangle(fid, 'Port', 'X', [-Lx/2 - Lfeed, -PortW/2, 0], PortW, -PortH, 'meter');
hfssAssignLumpedPort(fid, 'LPort', 'Port', [-Lx/2-Lfeed, 0, 0], ...
                     [-Lx/2-Lfeed, 0, -PortH], 'meter');

% Genera el substrato.
hfssBox(fid, 'Substrate', [-WsubX/2, -WsubY/2, 0], [WsubX, WsubY, -Hsub], 'meter');
hfssAssignMaterial(fid, 'Substrate', 'Rogers RT/duroid 5880 (tm)'); % Material.
hfssSetColor(fid, 'Substrate', [0, 128, 0]); % Color.
hfssSetTransparency(fid, {'Substrate'}, 0.2); % Transparencia.

% Clona los parches en el eje X...
if (Nx > 1)
	hfssMove(fid, {'Patch', 'Port'}, [-0.5*(Nx-1)*dAnt, 0, 0], 'meter');
	hfssDuplicateAlongLine(fid, {'Patch', 'Port'}, [dAnt, 0, 0], Nx, 'meter');
end;

% ... y termina el array clonando los parches en Y.
xObjList_patch = cell(Ny, 1); % Columna de 3x1.
xObjList_patch{1} = 'Patch'; % Nombre.
xObjList_port = cell(Ny, 1); % Columna de 3x1.
xObjList_port{1} = 'Port'; % Nombre.
for iY = 2:Ny,
	xObjList_patch{iY} = ['Patch_', num2str(iY-1)];
	xObjList_port{iY}  = ['Port_', num2str(iY-1)];
end;
xObjList = [xObjList_patch; xObjList_port];
if (Nx > 1)
	hfssMove(fid, xObjList , [0, -0.5*(Ny-1)*dAnt, 0], 'meter');
	hfssDuplicateAlongLine(fid, xObjList, [0, dAnt, 0], Ny, 'meter');
end;

% Crea la caja de radiaci�n.
hfssBox(fid, 'AirBox', [-AirX/2, -AirY/2, -Hsub], [AirX, AirY, AirZ], 'meter'); 
hfssAssignRadiation(fid, 'ABC', 'AirBox'); % La asigna como l�mite.
hfssSetTransparency(fid, {'AirBox'}, 0.95); % Transparencia.

% Crea el plano de tierra.
hfssRectangle(fid, 'GroundPlane', 'Z', [-AirX/2, -AirY/2, -Hsub], AirX, AirY, 'meter');
hfssAssignPE(fid, 'GND', {'GroundPlane'}); % Lo asigna como GND.

% Asigna una soluci�n y un barrido.
hfssInsertSolution(fid, 'Setup3_75GHz', fC/1e9); % Nombre y frecuencia.
hfssInterpolatingSweep(fid, 'Sweep3to4_5GHz', 'Setup3_75GHz', fStart, fStop, 1001); % Barrido.
hfssSaveProject(fid, 'C:\Users\Asus\Documents\Ansoft\antenaMatlab.aedt', true); % Guarda el HFSS (CAMBIAR).

% Cierra el objeto usado en la configuraci�n del script.
fclose(fid);

% Abre HFSS y ejecuta el script.
hfssExecuteScript(hfssExePath, tmpScriptFile, false, false);

