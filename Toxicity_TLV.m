% This code is developed by Kazi Monzure Khoda
% This code helps to find the coefficients of exponential equations for
% Toxicity using TLV-STEL
clc
clear
clear global
% clear all
close all

%-------------------------
% Curve fitting for Room temperature
%-------------------------
filename = 'Matlab Input.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder


sheet = 'Toxicity'; % This is the excel sheet name
%xlRange = 'C6:C32'; % This is the range of the cells in exel
xlRange = 'C3:C53'; % This is the range of the cells in exel

subsetA = xlsread(filename,sheet,xlRange);% This function call inputs from excel

sheet = 'Toxicity'; % This is the excel sheet name
%xlRange = 'D6:D32'; % This is the range of the cells in exel
xlRange = 'D3:D53'; % This is the range of the cells in exel

subsetB = xlsread(filename,sheet,xlRange);% This function call inputs from excel

[rows, columns] = size(subsetA); % This measure the size of the inputs
length = rows; 

for input = 1:length 

    for i = 1:length
x1(input) = subsetA(input);
y1(input)= 100-subsetB(input);
    end
end
x1=x1';
y1=y1';
g = fittype('(a*exp(b)^x)');
f1 = fit(x1,y1,g,'StartPoint',[0,0])
plot(f1,'r',x1,y1,'or')
ylim([0 100])
coeffvals1 = coeffvalues(f1);
hold on;
title('Exponential curve fitting for toxicity')
xlabel('TLV-STEL (ppm)') % x-axis label
ylabel('Scale') % y-axis label
legend('Data-Room Temp','Fitted Curve-Room Temp','Data-Op.Temp.1','Fitted Curve-Op.Temp.1','Data-Op.Temp.2','Fitted Curve-Op.Temp.2')
xlswrite('Matlab Input.xlsx',coeffvals3,'Toxicity','H8:I8')

