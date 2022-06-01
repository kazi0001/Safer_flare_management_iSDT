% This code is developed by Kazi Monzure Khoda
% This code helps to find the coefficients of exponential equations for
% different UEL%-LEL% at different temperature
clc
clear
clear global
% clear all
close all

%-------------------------
% Curve fitting for LEL25
%-------------------------
filename = 'Matlab Input.xlsx'; % Matlab will call this excel file for input, this file should be in the same folder


sheet = 'Explosiveness'; % This is the excel sheet name
xlRange = 'F3:F100'; % This is the range of the cells in exel

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
g = fittype('exp2');
f1 = fit(x1,y1,g)
plot(f1,'r',x1,y1,'or')
coeffvals1 = coeffvalues(f1);
hold on;
xlswrite('Matlab Input.xlsx',coeffvals1,'Explosiveness','K7:L7')
% 
% %-------------------------------------------------
% 
% %-------------------------
% % Curve fitting for LEL500
% %-------------------------
% 
% sheet = 'Explosiveness';
% xlRange = 'G3:G100';
% 
% subsetB = xlsread(filename,sheet,xlRange);
% 
% [rows, columns] = size(subsetB);
% length = rows;
% 
% for input = 1:length 
% 
%     for i = 1:length
% x2(input) = (input);
% y2(input)= subsetB(input);
%     end
% end
% x2=x2';
% y2=y2';
% g = fittype('(a*exp(b)^x)+ c*exp(d)^x)');
% f2 = fit(x2,y2,g,'StartPoint',[0,0])
% plot(f2,'g',x2,y2,'xg')
% coeffvals2 = coeffvalues(f2);
% hold on;
% xlswrite('Matlab Input.xlsx',coeffvals2,'Explosiveness','K10:L10')
% %--------------------------------------
% 
% %-------------------------
% % Curve fitting for LEL1000
% %-------------------------
% 
% sheet = 'Explosiveness';
% xlRange = 'H3:H100';
% 
% subsetC = xlsread(filename,sheet,xlRange);
% 
% [rows, columns] = size(subsetC);
% length = rows;
% 
% for input = 1:length 
% 
%     for i = 1:length
% x3(input) = (input);
% y3(input)= subsetC(input);
%     end
% end
% x3=x3';
% y3=y3';
% g = fittype('(a*exp(b)^x)+ c*exp(d)^x)');
% f3 = fit(x3,y3,g,'StartPoint',[0,0])
% plot(f3,'b',x3,y3,'*b')
% ylim([0 100])
% coeffvals3 = coeffvalues(f3);
% title('Exponential curve fitting for explosiveness')
% xlabel('UEL%-LEL%') % x-axis label
% ylabel('Scale') % y-axis label
% legend('Data-LEL25','Fitted Curve-LEL25','Data-LEL500','Fitted Curve-LEL500','Data-LEL1000','Fitted Curve-LEL1000')
% hold on;
% xlswrite('Matlab Input.xlsx',coeffvals1,'Explosiveness','K14:L14')