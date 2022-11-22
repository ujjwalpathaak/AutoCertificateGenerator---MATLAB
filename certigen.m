clc 
clear all
close all 
%home 

filename = 'database.xlsx';
[num,txt] = xlsread(filename);
len=length(txt);
blankimg = imread('certificate_template.png');

for i=1:len
        text_names(i,2)=txt(i,1);
end

% Obtain names from the txt variable which are in 2nd column

for i=1:len
      text_contribution(i,3)=txt(i,2);
end

% Obtain contribution from the txt variable which are in 3nd column

for i=1:len
        text_event(i,4)=txt(i,3);
end

% Obtain event from the txt variable which are in 4th column

for i=1:len
      text_date(i,5)=txt(i,4);
end

% Obtain date from the txt variable which are in 5th column

for i=2:len
        text_all=[text_names(i,2) text_topic(i,3) text_event(i,4) text_date(i,5)];
        
        position = [190 830;825 990;190 1060;1125 1230];       
        
        RGB = insertText(blankimg,position,text_all,'FontSize',82,'BoxOpacity',0);
        %Default font with Font size is 22 and opacity of box is 0 means 100% transparent
        
        figure
        imshow(RGB);       
        %Obtain and display figure in color
        
        y=i-1;
        filename=['CertiJR' num2str(y)];
        lastf=[filename '.png'];
        saveas(gcf,lastf);
        % generate and save files with same extension as original img
end    
