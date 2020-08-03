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
fC = 2.4e9; % CAMBIAR a su frecuencia de funcionamiento.
Lambda = c/fC;

% Dimensiones del parche.
Wx = 24*1e-3;
Ly = 24*1e-3;

radio = 17.4*1e-3;

% Par�metros del array.
Nx = 1;
Ny = 1;
dAnt = Lambda/2;

% Dimensiones de la alimentaci�n.
Wms = 11*1e-3;
Lfeed = radio;

% Par�metros del substrato.
Hsub = 1.66*1e-3;
WsubX = Nx*55*1e-3;
WsubY = Ny*50*1e-3;

% Par�metros del puerto.
PortH = Hsub;
PortW = Wms;

% Par�metros caja de radiaci�n.
AirX = Nx*dAnt + Lambda/2;
AirY = Ny*dAnt + Lambda/2;
AirZ = Lambda/4 + Hsub;
	
% Par�metros de soluci�n (Solution Setup) (GHz).
fSolve = fC/1e9;
fStart = 1;
fStop  = 5;

% Abre un script temporal y almacena las �rdenes.
tmpScriptFile = 'tempPatch.vbs';
fid = fopen(tmpScriptFile, 'wt');

% Crea un nuevo proyecto en HFSS.
hfssNewProject(fid);
hfssInsertDesign(fid, 'ParcheCiruclarSolo'); % Nombre.

% Genera el parche y la alimentaci�n.
hfssCircle(fid, 'Patch', 'Z', [0, 0, 0], radio, 'meter');
hfssRectangle(fid, 'Feed', 'Z', [-WsubX/2, -Wms/2, 0],  Lfeed, Wms, 'meter');
% Une el parche y la alimentaci�n
hfssUnite(fid, {'Patch', 'Feed'});
% Asigna l�mite al parche (Perfect E).
hfssAssignPE(fid, 'PatchMetal', {'Patch'});
hfssSetColor(fid, 'Patch', [128, 128, 0]); % Color.
hfssSetTransparency(fid, {'Patch'}, 0); % Transparencia.

% Genera un lumped port y lo asigna.
hfssRectangle(fid, 'Port', 'X', [-WsubX/2, -Wms/2, 0], PortW, -PortH, 'meter');
hfssAssignLumpedPort(fid, 'LPort', 'Port', [-WsubX/2, 0, 0], ...
                      [-WsubX/2, 0, -PortH], 'meter');

% Genera el substrato.
hfssBox(fid, 'Substrate', [-WsubX/2, -WsubY/2, 0], [WsubX, WsubY, -Hsub], 'meter');
hfssAssignMaterial(fid, 'Substrate', 'FR4_epoxy'); % Material.
hfssSetColor(fid, 'Substrate', [0, 128, 0]); % Color.
hfssSetTransparency(fid, {'Substrate'}, 0.2); % Transparencia.

% Crea la caja de radiaci�n.
hfssBox(fid, 'AirBox', [-WsubX/2, -WsubY/2, -1.5*Hsub], [WsubX, WsubY, AirZ], 'meter'); 
hfssAssignRadiation(fid, 'Radiacion', 'AirBox'); % La asigna como l�mite.
hfssSetTransparency(fid, {'AirBox'}, 0.95); % Transparencia.

% Crea el plano de tierra.
hfssRectangle(fid, 'PlanoTierra', 'Z', [-WsubX/2, -WsubY/2, -Hsub], WsubX, WsubY, 'meter');
hfssAssignPE(fid, 'GND', {'PlanoTierra'}); % Lo asigna como GND.

% Asigna una soluci�n y un barrido.
hfssInsertSolution(fid, 'Setup2_4GHz', fC/1e9); % Nombre y frecuencia.
hfssInterpolatingSweep(fid, 'Sweepto1_5GHz', 'Setup2_4GHz', fStart, fStop, 201); % Barrido.
hfssSaveProject(fid, 'C:\Users\Asus\Documents\Ansoft\antenaEjemploMatlab.aedt', true); % Guarda el HFSS (CAMBIAR).

% Cierra el objeto usado en la configuraci�n del script.
fclose(fid);

% Abre HFSS y ejecuta el script.
hfssExecuteScript(hfssExePath, tmpScriptFile, false, false);
