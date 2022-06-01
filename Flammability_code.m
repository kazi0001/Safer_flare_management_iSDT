clc
clear
clear global
% clear all
close all

filename = 'Test Matlab.xlsx';
sheet = 7;
xlRange = 'D4:D104';


subsetA = xlsread(filename,sheet,xlRange);

[rows, columns] = size(subsetA);
length = rows;

for input = 1:length 

    for i = 1:length
x(input) = 200+subsetA(input);
y(input)= 100-(input);
    end
end
x=x';
y=y';
f = fit(x,y,'exp1','StartPoint',[0,0],'Upper',100)

plot(f,x,y)
set(gca,'XTick',[0 100 200 300 400 500 600] );
set(gca,'XTickLabel',[-200 -100 0 100 200 300 400] );

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