clc
clear
clear global
% clear all
close all

filename = 'Test Matlab.xlsx';
sheet = 5;
xlRange = 'D4:D54';


subsetA = xlsread(filename,sheet,xlRange);
subsetA(1)
[rows, columns] = size(subsetA);
length = rows;

for input = 1:length 

    for i = 1:length
x(input) = 22+(input);
y(input)= subsetA(input);
    end
end
x=x';
y=y';
f = fit(x,y,'exp1','StartPoint',[0,0],'Upper',100)
plot(f,x,y)


% x = 0:0.1:100;
% y = 100*(exp((-.05*x)));
% figure % opens new figure window
% plot(x,y)
% hold on;
% 
% y = .67*(exp((.05*x)));
% 
% figure % opens new figure window
% plot(x,y)