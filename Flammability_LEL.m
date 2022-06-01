% This code is developed by Kazi Monzure Khoda
% This code helps to find the coefficients of exponential equations for
% different LEL at different temperature
clc
clear
clear global
% clear all
close all

%-------------------------
% Curve fitting for Flare Locatoin A
%-------------------------
filename = 'Matlab Input Flare source.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder


sheet = 'Flammability_LFL'; % This is the excel sheet name
xlRange = 'D10:D60'; % This is the range of the cells in exel

subsetA = xlsread(filename,sheet,xlRange);% This function call inputs from excel

[rows, columns] = size(subsetA); % This measure the size of the inputs
length = rows; 

for input = 1:length 

    for i = 1:length
x1(input) = (input);
y1(input)= 100-subsetA(input);
    end
end
x1=x1';
y1=y1';
g = fittype('(100-a*exp(b)^x)');
f1 = fit(x1,y1,g,'StartPoint',[0,0])
plot(f1,'r',x1,y1,'or')
coeffvals1 = coeffvalues(f1);
hold on;
xlswrite('Matlab Input Flare source.xlsx',coeffvals1,'Flammability_LFL','O15:P15')

%-------------------------------------------------

%-------------------------
% Curve fitting for Flare Locatoin B
%-------------------------
filename = 'Matlab Input Flare source.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder


sheet = 'Flammability_LFL'; % This is the excel sheet name
xlRange = 'E10:E60'; % This is the range of the cells in exel

subsetB = xlsread(filename,sheet,xlRange);% This function call inputs from excel

[rows, columns] = size(subsetB); % This measure the size of the inputs
length = rows; 

for input = 1:length 

    for i = 1:length
x2(input) = (input);
y2(input)= 100-subsetB(input);
    end
end
x2=x2';
y2=y2';
g = fittype('(100-a*exp(b)^x)');
f2 = fit(x2,y2,g,'StartPoint',[0,0])
plot(f2,'b',x2,y2,'ob')
coeffvals2 = coeffvalues(f2);
hold on;
xlswrite('Matlab Input Flare source.xlsx',coeffvals2,'Flammability_LFL','O17:P17')

%-------------------------------------------------
%-------------------------
% Curve fitting for Flare Locatoin C
%-------------------------
filename = 'Matlab Input Flare source.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder


sheet = 'Flammability_LFL'; % This is the excel sheet name
xlRange = 'F10:F60'; % This is the range of the cells in exel

subsetC = xlsread(filename,sheet,xlRange);% This function call inputs from excel

[rows, columns] = size(subsetC); % This measure the size of the inputs
length = rows; 

for input = 1:length 

    for i = 1:length
x3(input) = (input);
y3(input)= 100-subsetC(input);
    end
end
x3=x3';
y3=y3';
g = fittype('(100-a*exp(b)^x)');
f3 = fit(x3,y3,g,'StartPoint',[0,0])
plot(f3,'k',x3,y3,'ok')
coeffvals3 = coeffvalues(f3);
hold on;
xlswrite('Matlab Input Flare source.xlsx',coeffvals3,'Flammability_LFL','O19:P19')

%-------------------------
% Curve fitting for Flare Locatoin D
%-------------------------
filename = 'Matlab Input Flare source.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder


sheet = 'Flammability_LFL'; % This is the excel sheet name
xlRange = 'G10:G60'; % This is the range of the cells in exel

subsetD = xlsread(filename,sheet,xlRange);% This function call inputs from excel

[rows, columns] = size(subsetD); % This measure the size of the inputs
length = rows; 

for input = 1:length 

    for i = 1:length
x4(input) = (input);
y4(input)= 100-subsetD(input);
    end
end
x4=x4';
y4=y4';
g = fittype('(100-a*exp(b)^x)');
f4 = fit(x4,y4,g,'StartPoint',[0,0])
plot(f4,'g',x4,y4,'og')
coeffvals4 = coeffvalues(f4);
hold on;
xlswrite('Matlab Input Flare source.xlsx',coeffvals4,'Flammability_LFL','O21:P21')

%-------------------------
%-------------------------
% Curve fitting for Flare Locatoin E
%-------------------------
filename = 'Matlab Input Flare source.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder


sheet = 'Flammability_LFL'; % This is the excel sheet name
xlRange = 'H10:H60'; % This is the range of the cells in exel

subsetE = xlsread(filename,sheet,xlRange);% This function call inputs from excel

[rows, columns] = size(subsetE); % This measure the size of the inputs
length = rows; 

for input = 1:length 

    for i = 1:length
x5(input) = (input);
y5(input)= 100-subsetE(input);
    end
end
x5=x5';
y5=y5';
g = fittype('(100-a*exp(b)^x)');
f5 = fit(x5,y5,g,'StartPoint',[0,0])
plot(f5,'y',x5,y5,'oy')
coeffvals5 = coeffvalues(f5);
hold on;
xlswrite('Matlab Input Flare source.xlsx',coeffvals5,'Flammability_LFL','O23:P23')

%-------------------------
%-------------------------
% Curve fitting for Flare Locatoin F
%-------------------------
filename = 'Matlab Input Flare source.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder


sheet = 'Flammability_LFL'; % This is the excel sheet name
xlRange = 'I10:I60'; % This is the range of the cells in exel

subsetF = xlsread(filename,sheet,xlRange);% This function call inputs from excel

[rows, columns] = size(subsetF); % This measure the size of the inputs
length = rows; 

for input = 1:length 

    for i = 1:length
x6(input) = (input);
y6(input)= 100-subsetF(input);
    end
end
x6=x6';
y6=y6';
g = fittype('(100-a*exp(b)^x)');
f6 = fit(x6,y6,g,'StartPoint',[0,0])
plot(f6,'c',x6,y6,'oc')
coeffvals6 = coeffvalues(f6);
hold on;
xlswrite('Matlab Input Flare source.xlsx',coeffvals6,'Flammability_LFL','O25:P25')

%-------------------------

%-------------------------
% Curve fitting for Flare Locatoin G
%-------------------------
filename = 'Matlab Input Flare source.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder

sheet = 'Flammability_LFL';
xlRange = 'J10:J60';

subsetG = xlsread(filename,sheet,xlRange);

[rows, columns] = size(subsetG);
length = rows;

for input = 1:length 

    for i = 1:length
x7(input) = (input);
y7(input)= 100-subsetG(input);
    end
end
x7=x7';
y7=y7';
g = fittype('(100-a*exp(b)^x)');
f7 = fit(x7,y7,g,'StartPoint',[0,0])
plot(f7,'m',x7,y7,'*m')
coeffvals7 = coeffvalues(f7);
title('Exponential curve fitting for flammability')
xlabel('LEL%') % x-axis label
ylabel('Scale') % y-axis label
legend('Data-Room Temp','Fitted Curve-Room Temp','Data-Op.Temp.1','Fitted Curve-Op.Temp.1','Data-Op.Temp.2','Fitted Curve-Op.Temp.2')
hold on;
xlswrite('Matlab Input Fadwa.xlsx',coeffvals7,'Flammability_LFL','O27:P27')