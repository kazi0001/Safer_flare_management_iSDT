% This code is developed by Kazi Monzure Khoda
% This code helps to find the coefficients of exponential equations for
% different UEL%-LEL% at different temperature
clc
clear
clear global
% clear all
close all

%-------------------------
% Curve fitting for Location A
%-------------------------
filename = 'Matlab Input Flare source.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder


sheet = 'Explosiveness'; % This is the excel sheet name
xlRange = 'F11:F116'; % This is the range of the cells in exel

subsetA = xlsread(filename,sheet,xlRange);% This function call inputs from excel

[rows, columns] = size(subsetA); % This measure the size of the inputs
length = rows; 

for input = 1:length 

    for i = 1:length
x1(input) = (input);
y1(input)= subsetA(input);
    end
end
x1=x1';
y1=y1';
g = fittype('(a*exp(b)^x)');
f1 = fit(x1,y1,g,'StartPoint',[0,0])
plot(f1,'r',x1,y1,'or')
coeffvals1 = coeffvalues(f1);
hold on;
xlswrite('Matlab Input Flare source.xlsx',coeffvals1,'Explosiveness','O16:P16')

%-------------------------------------------------
%-------------------------
% Curve fitting for Location B
%-------------------------
filename = 'Matlab Input Flare source.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder


sheet = 'Explosiveness'; % This is the excel sheet name
xlRange = 'G11:G116'; % This is the range of the cells in exel

subsetB = xlsread(filename,sheet,xlRange);% This function call inputs from excel

[rows, columns] = size(subsetB); % This measure the size of the inputs
length = rows; 

for input = 1:length 

    for i = 1:length
x2(input) = (input);
y2(input)= subsetB(input);
    end
end
x2=x2';
y2=y2';
g = fittype('(a*exp(b)^x)');
f2 = fit(x2,y2,g,'StartPoint',[0,0])
plot(f2,'b',x2,y2,'ob')
coeffvals2 = coeffvalues(f2);
hold on;
xlswrite('Matlab Input Flare source.xlsx',coeffvals2,'Explosiveness','O18:P18')

%-------------------------------------------------
%-------------------------
% Curve fitting for Location C
%-------------------------
filename = 'Matlab Input Flare source.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder


sheet = 'Explosiveness'; % This is the excel sheet name
xlRange = 'H11:H116'; % This is the range of the cells in exel

subsetC = xlsread(filename,sheet,xlRange);% This function call inputs from excel

[rows, columns] = size(subsetC); % This measure the size of the inputs
length = rows; 

for input = 1:length 

    for i = 1:length
x3(input) = (input);
y3(input)= subsetC(input);
    end
end
x3=x3';
y3=y3';
g = fittype('(a*exp(b)^x)');
f3 = fit(x3,y3,g,'StartPoint',[0,0])
plot(f3,'k',x3,y3,'ok')
coeffvals3 = coeffvalues(f3);
hold on;
xlswrite('Matlab Input Flare source.xlsx',coeffvals3,'Explosiveness','O20:P20')

%-------------------------------------------------%-------------------------
% Curve fitting for Location D
%-------------------------
filename = 'Matlab Input Flare source.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder


sheet = 'Explosiveness'; % This is the excel sheet name
xlRange = 'I11:I116'; % This is the range of the cells in exel

subsetD = xlsread(filename,sheet,xlRange);% This function call inputs from excel

[rows, columns] = size(subsetD); % This measure the size of the inputs
length = rows; 

for input = 1:length 

    for i = 1:length
x4(input) = (input);
y4(input)= subsetD(input);
    end
end
x4=x4';
y4=y4';
g = fittype('(a*exp(b)^x)');
f4 = fit(x4,y4,g,'StartPoint',[0,0])
plot(f4,'g',x4,y4,'og')
coeffvals4 = coeffvalues(f4);
hold on;
xlswrite('Matlab Input Flare source.xlsx',coeffvals4,'Explosiveness','O22:P22')

%-------------------------------------------------%-------------------------
% Curve fitting for Location E
%-------------------------
filename = 'Matlab Input Flare source.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder


sheet = 'Explosiveness'; % This is the excel sheet name
xlRange = 'J11:J116'; % This is the range of the cells in exel

subsetE = xlsread(filename,sheet,xlRange);% This function call inputs from excel

[rows, columns] = size(subsetE); % This measure the size of the inputs
length = rows; 

for input = 1:length 

    for i = 1:length
x5(input) = (input);
y5(input)= subsetA(input);
    end
end
x5=x5';
y5=y5';
g = fittype('(a*exp(b)^x)');
f5 = fit(x5,y5,g,'StartPoint',[0,0])
plot(f5,'y',x5,y5,'oy')
coeffvals5 = coeffvalues(f5);
hold on;
xlswrite('Matlab Input Flare source.xlsx',coeffvals5,'Explosiveness','O24:P24')

%-------------------------------------------------%-------------------------
% Curve fitting for Location F
%-------------------------
filename = 'Matlab Input Flare source.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder


sheet = 'Explosiveness'; % This is the excel sheet name
xlRange = 'K11:K116'; % This is the range of the cells in exel

subsetF = xlsread(filename,sheet,xlRange);% This function call inputs from excel

[rows, columns] = size(subsetF); % This measure the size of the inputs
length = rows; 

for input = 1:length 

    for i = 1:length
x6(input) = (input);
y6(input)= subsetF(input);
    end
end
x6=x6';
y6=y6';
g = fittype('(a*exp(b)^x)');
f6 = fit(x6,y6,g,'StartPoint',[0,0])
plot(f6,'c',x6,y6,'oc')
coeffvals6 = coeffvalues(f6);
hold on;
xlswrite('Matlab Input Flare source.xlsx',coeffvals6,'Explosiveness','O26:P26')

%-------------------------------------------------
%-------------------------------------------------%-------------------------
% Curve fitting for Location G
%-------------------------
filename = 'Matlab Input Flare source.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder

sheet = 'Explosiveness';
xlRange = 'L11:L116';

subsetG = xlsread(filename,sheet,xlRange);

[rows, columns] = size(subsetG);
length = rows;

for input = 1:length 

    for i = 1:length
x7(input) = (input);
y7(input)= subsetG(input);
    end
end
x7=x7';
y7=y7';
g = fittype('(a*exp(b)^x)');
f7 = fit(x7,y7,g,'StartPoint',[0,0])
plot(f7,'m',x7,y7,'*m')
ylim([0 100])
coeffvals7 = coeffvalues(f7);
title('Exponential curve fitting for explosiveness')
xlabel('UEL%-LEL%') % x-axis label
ylabel('Scale') % y-axis label
legend('Data-Room Temp','Fitted Curve-Room Temp','Data-Op.Temp.1','Fitted Curve-Op.Temp.1','Data-Op.Temp.2','Fitted Curve-Op.Temp.2')
hold on;
xlswrite('Matlab Input Fadwa.xlsx',coeffvals7,'Explosiveness','O28:P28')