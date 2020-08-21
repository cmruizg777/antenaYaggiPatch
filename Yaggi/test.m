%% PAPER GUÍA: https://ieeexplore.ieee.org/stamp/stamp.jsp?tp=&arnumber=8454719


%% Limpiar el workspace y la ventana de comandos.
fclose all;
clear all;
close all;
clc;

%% Agregar al path las funciones del API.

addpath('../HFSS_API-master/boundary/');
addpath('../HFSS_API-master/3dmodeler/');
addpath('../HFSS_API-master/analysis/');
addpath('../HFSS_API-master/general/');
addpath('../HFSS_API-master/reporter/');

%% Datos del proyecto

hfssExePath = '"C:\Program Files\AnsysEM\AnsysEM19.2\Win64\ansysedt.exe"';
tmpScriptFile = 'yaggi.vbs';
fid = fopen(tmpScriptFile, 'wt');
hfssNewProject(fid);
hfssInsertDesign(fid, 'yaggi'); 

%% VARIBLES
f = 24e6;
DirSpa = 3;

Lt1  = 1.64;
Lt2  = 1.91;
Ldir = 2.75;
Ldip = 5.19;
Lcps = 1;

Wt1  = 0.65;
Wt2  = 0.54;
Wdir = 0.72;
Wdip = 0.8;
Wcps = 1;
Wmsc = 1;

gap = 0.3;
s = 0.5;


Lt0 = 2.5;
W = 20;
Wt0 = 1;

material = 'FR4_epoxy';
units = 'mm';
% el ancho del puerto
portW = Wt0;

%% FEED

hfssRectangle(fid, 'r1', 'Z', [0 -Wt0/2 0] ,  Lt0, Wt0, units);
 
%% Line 2

hfssRectangle(fid, 'r2', 'Z', [Lt0 -Wt2/2 0] ,  Lt2, Wt2, units);
 
%% Line 1

hfssRectangle(fid, 'r3', 'Z', [Lt0+Lt2 -Wt1/2 0] ,  Lt1, Wt1, units);

%% transformador

hfssRectangle(fid, 'r4', 'Z', [Lt0+Lt1+Lt2 -Ldip/2 0] ,  3*Wmsc, Ldip, units);
Lgnd = Lt0+Lt1+Lt2+3*Wmsc;

%% Separación Lcps*(Wcps+gap)

w = Wcps+gap;
hfssRectangle(fid, 'r5', 'Z', [Lt0+Lt1+Lt2+3*Wmsc -w/2 0] ,  Lcps, w, units);
%% Dipolo Wdip*Ldip

hfssRectangle(fid, 'r6', 'Z', [Lt0+Lt1+Lt2+Lcps+3*Wmsc -Ldip/2 0] ,  Wdip, Ldip, units);

%% Director1 Wdir*Ldir

hfssRectangle(fid, 'r7', 'Z', [Lt0+Lt1+Lt2+Lcps+Wdip+DirSpa+3*Wmsc -Ldir/2 0] ,  Wdir, Ldir , units);

%% Director2 Wdir*Ldir

hfssRectangle(fid, 'r8', 'Z', [Lt0+Lt1+Lt2+Lcps+2*Wdip+2*DirSpa+3*Wmsc -Ldir/2 0] ,  Wdir, Ldir , units);
%% Director3 Wdir*Ldir

hfssRectangle(fid, 'r9', 'Z', [Lt0+Lt1+Lt2+Lcps+3*Wdip+3*DirSpa+3*Wmsc -Ldir/2 0] ,  Wdir, Ldir , units);

%% Dielectrico -> sustrato
L = Lt0+Lt1+Lt2+Lcps+4*Wdip+3*DirSpa+3*Wmsc+gap;
hfssBox(fid, 'Sustrato', [0 , -W/2 , 0], [L , W, -s ],units);
hfssAssignMaterial(fid,'Sustrato',material);

%% Unir segmentos patch
names = {};
for i=1:9
    name = strcat('r',num2str(i));
    names(i)={name}; 
end
hfssUnite(fid,names);
hfssSetColor(fid, 'r1', [220, 220, 10]);
hfssSetTransparency(fid, {'r1'}, 0);    
hfssAssignFiniteCond(fid, 'Patch', 'r1', 'false',false);

%% extract1 (Lcps + Wdip)*gap

% rectangulo divisor
hfssRectangle(fid, 'e1', 'Z', [Lt0+Lt1+Lt2+1.5*Wmsc -gap/2 0] ,  Lcps+Wdip+1.5*Wmsc, gap, units);
hfssSetColor(fid, 'e1', [220, 0, 0]);
% rombo
x1 = [Lt0+Lt1+Lt2+3*Wmsc; Lt0+Lt1+Lt2+1.5*Wmsc; Lt0+Lt1+Lt2; Lt0+Lt1+Lt2+1.5*Wmsc; Lt0+Lt1+Lt2+3*Wmsc];
y1 = [0; -1.5*Wmsc; 0; 1.5*Wmsc; 0];
puntos = [x1,y1,zeros(length(x1),1)];
hfssPolyline(fid,'e2' , puntos, 'mm', 'true' );
hfssSetColor(fid, 'e2', [220, 0, 0]);
% rectangulo paralelo al dipolo
hfssRectangle(fid, 'e3', 'Z', [Lt0+Lt1+Lt2+1*Wmsc Wmsc-Ldip/2 0] ,  Wmsc, Ldip -2*Wmsc, units);
hfssSetColor(fid, 'e3', [220, 0, 0]);

% esquinas (triangulos)
x2 = [Lt0+Lt1+Lt2+3*Wmsc; Lt0+Lt1+Lt2+3*Wmsc; Lt0+Lt1+Lt2+2*Wmsc; Lt0+Lt1+Lt2+3*Wmsc];
y2 = [-Ldip/2;Wmsc-Ldip/2; -Ldip/2; -Ldip/2];
puntos2 = [x2,y2,zeros(length(x2),1)];
hfssPolyline(fid,'e4' , puntos2, 'mm', 'true' );

x3 = [Lt0+Lt1+Lt2+3*Wmsc; Lt0+Lt1+Lt2+3*Wmsc; Lt0+Lt1+Lt2+2*Wmsc; Lt0+Lt1+Lt2+3*Wmsc];
y3 = [Ldip/2;-Wmsc+Ldip/2; Ldip/2; Ldip/2];
puntos3 = [x3,y3,zeros(length(x3),1)];
hfssPolyline(fid,'e5' , puntos3, 'mm', 'true' );

x4 = [Lt0+Lt1+Lt2; Lt0+Lt1+Lt2+Wmsc; Lt0+Lt1+Lt2; Lt0+Lt1+Lt2];
y4 = [-Ldip/2;-Ldip/2; Wmsc-Ldip/2; -Ldip/2];
puntos4 = [x4,y4,zeros(length(x4),1)];
hfssPolyline(fid,'e6' , puntos4, 'mm', 'true' );

x5 = [Lt0+Lt1+Lt2; Lt0+Lt1+Lt2+Wmsc; Lt0+Lt1+Lt2; Lt0+Lt1+Lt2];
y5 = [Ldip/2;Ldip/2; -Wmsc+Ldip/2; Ldip/2];
puntos5 = [x5,y5,zeros(length(x5),1)];
hfssPolyline(fid,'e7' , puntos5, 'mm', 'true' );

%hfssRectangle(fid, 'e1', 'Z', [Lt0+Lt1+Lt2+1.5*Wmsc -gap/2 0] ,  Lcps+Wdip+1.5*Wmsc, gap, units);


names2 = {};
for i=1:7
    name = strcat('e',num2str(i));
    names2(i)={name}; 
end

hfssSubtract(fid, {'r1'}, names2, false);

%% GND
hfssRectangle(fid, 'gnd', 'Z', [0 -W/2 -s] ,  Lgnd, W , units);
hfssSetColor(fid, 'gnd', [10, 220, 10]);
hfssSetTransparency(fid, {'gnd'}, 0);    
hfssAssignPE(fid, 'GND', {'gnd'}, 'false');


%% PORT

hfssRectangle(fid, 'port', 'X', [0 -Wt0/2 -s] , Wt0 , s , units);
hfssSetColor(fid, 'port', [235 10 10]);
hfssSetTransparency(fid, {'port'}, 0);
hfssAssignLumpedPort(fid, 'puerto', 'port', [0, 0, -s/2], [0, 0, 0], units, 50, 0);

%% Radiation

hfssBox(fid, 'box2', [-2*s , -2*s-W/2 , -2*s ], [L + 4*s, W+4*s, s*20 ],units);
hfssSetTransparency(fid, {'box2'}, 0.95);
hfssAssignRadiation(fid, 'rad1', 'box2')

%% insertar solución, crear reporte y guardar el proyecto

hfssInsertSolution(fid, 'yaggi_uda', 5);
hfssInterpolatingSweep(fid, 'SWEEP1', 'yaggi_uda', 3, 7, 101);
hfssCreateReport(fid, 'Report 1', 1, 1, 'yaggi_uda', 'Sweep1',...
                       [], 'Sweep', {'Freq'}, {'Freq',...
                       'db(S(puerto,puerto))'});
directory_content = dir; % contains everything of the current directory
path = directory_content(1).folder; % returns the path that is currently open
hfssSaveProject(fid, strcat(path,'\',tmpScriptFile,'.aedt'), true); % Guarda el HFSS (CAMBIAR).
% hfssSolveSetup(fid, 'nfc');

%% Cierra el objeto usado en la configuración del script.
fclose(fid);
% %Abre HFSS y ejecuta el script.
hfssExecuteScript(hfssExePath, tmpScriptFile, false, false);