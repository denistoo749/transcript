warning('off','all')

%Read data from excel
result = readtable('MCS_3.1.xlsx');
%studentinfo = table2array(result(:,[1,2]));
studentinfo = table2array(readtable('MCS_3.1.xlsx','Range','A2:B169'));
scores = table2array(readtable('MCS_3.1.xlsx','Range','C2:G169'));
units = string(result.Properties.VariableNames(3:end));
    
%input the reg number
reg = input('Enter the registration number for students:','s');
%prints transcript
fprintf("\tSTUDENT'S TRANSCRIPT\n")
printTranscript(scores,reg,units,studentinfo);
%print average and grading per unit
fprintf('\n\n')
fprintf("\t\tMEAN PER UNIT\n\n")
fprintf('%-30s%-20s%s\n', ' UNITS ',' AVERAGE ',' GRADE ');
for j=1:length(units)
    fprintf('%-30s%-30.2f%-15s \n',units(j),mean(scores(:,j)),grade(mean(scores(:,j))));
end

%print the table Result
fprintf('\n')
% Get the size of the table
[numRows, numCols] = size(result);

fprintf('%-20s', result.Properties.VariableNames{1});
for j = 2:numCols
    fprintf('%-30s', result.Properties.VariableNames{j});
end
fprintf('\n');

% Nested for loops to iterate through each element of the table
for i = 1:numRows
    for j = 1:numCols
        % Access the element at (i, j)
        element = result{i, j};
        
        % Print the element with fixed width
        if j == 1
            fprintf('%-20s', string(element)); % First column
        else
            fprintf('%-30s', string(element)); % Subsequent columns
        end
    end
    fprintf('\n');
end

function [] = printTranscript(scores,reg,units,studentinfo)
    for i=1:length(string(studentinfo(:,1)))
        if strcmp(string(studentinfo(i,1)), reg)
            fprintf('%s: %s\n\n%s: %s\n\n','Student Name',string(studentinfo(i,2)),'Registration Number',string(studentinfo(i,1)));
            fprintf('%-30s%-30s%-30s\n', 'UNITS','SCORES','GRADE');
            for k=1:length(units)
                fprintf('\n%-40s %-40d %-10s\n',units(k),scores(i,k),grade(scores(i,k)));
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