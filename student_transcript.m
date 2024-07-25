warning('off','all')
%Read data from excel
result = readtable('C:\Users\Denis\MCS_3.1.xlsx');
studentinfo = table2array(result(:,[1,2]));
scores = table2array(readtable('C:\Users\Denis/MCS_3.1.xlsx','Range','C2:I169'));
units = string(result.Properties.VariableNames(3:end));
%input the reg number
reg = input('Enter the registration number for students:','s');
%prints transcript
fprintf("\tSTUDENT'S TRANSCRIPT")
printTranscript(scores,reg,units,studentinfo);
%print average and grading per unit
fprintf("\t\tMEAN PER UNIT")
fprintf('%-30s%-20s%s\n', ' UNITS ',' AVERAGE ',' GRADE ');
for j=1:length(units)
    fprintf('%-30s%-30.2f%-15s \n',units(j),mean(scores(:,j)),grade(mean(scores(:,j))));
end

function [] = printTranscript(scores,reg,units,studentinfo)
    for i=1:length(string(studentinfo(:,1)))
        if string(studentinfo(i,1))==reg
            fprintf('%s: %s\n\n%s: %s\n\n','Student Name',string(studentinfo(i,2)),'Registration Number',string(studentinfo(i,1)));
            fprintf('%-30s%-30s\n', 'UNITS','GRADE');
            for k=1:length(units)
                fprintf('\n%-40s %-10s\n',units(k),grade(scores(i,k)));
            end
            break
        end
    end   
end
%grading the scores
function grd = grade(scores)
    if scores>=70
        grd = 'A';
    elseif scores>=60
        grd = 'B';
    elseif scores>=50
        grd = 'C';
    elseif scores>=40
        grd = 'D';
    else
        grd = 'E';
    end      
end