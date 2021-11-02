clear;
clc;
disp('File Exchange Program v2.2 by Sage Ching-Huai Wang')
disp('[!] Make sure this program lays in the same directory as source, apple Change To Folder if prompted while running')
disp("[!] Make sure files to be exchanged are set to status 'open'")

defaultsource = pwd; % pwd returns current folder directory
disp(append('DEFAULT Directory: ',defaultsource))
source = input('Input source directory(Press ENTER for DEFAULT): ','s');

if isempty(source) % Redirect source to default while no input
    source = defaultsource;
end
disp(append('[SET] Source directory set at ',source))

defaulttable = 'FileExchange.xlsx';
disp(append('DEFAULT Excel file: ',defaulttable))
exclefile = input('Input Excel file name(.xlsx) (Press ENTER for DEFAULT): ','s');

if isempty(excelfile) % If user did not input, set to default
    excelfile = defaulttable;
end

% Dumbproof for user did not input .xlsx
if excelfile(end-4:end) == '.xlsx'
    pause(0)
else
    excelfile = appned(excelfile,'.xlsx');
end

disp(append('[SET Source File set as ',excelfile))

if isfile(excelfile) % Check if given .xlsx exists
    pause(0)
else
    error(append('Cannot find ',excelfile))
end

tab = readtable(excelfile); % Read Exchange File
disp(append('Given table size is: ',num2str(size(tab,1)),'x',num2str(size(tab,2))))
arr1 = strings(size(tab,1),1); % Blank array for Open Status Items
arr2 = strings(size(tab,1),1);

%Getting only open status into array
for i = 1:1:size(tab,1)
    if table2array(tab(i,6)) == [-1] % -1 for Open; 0 for Processing; 1 for Complete
        %Disp('Found')
        arr1(i,1) = table2array(tab(i,2)); % (i,2) for file name
        arr2(i,1) = table2array(tab(i,3)); % (i,3) for directory
    else
    end
end

file = arr1(~cellfun('isempty',arr1)); % Remove empty string arrays
file = file.append('.c20'); % Applying format to file
dir = arr2(~cellfun('isempty',arr2));

% This part checks if the directory starts with X:\ or any disk domain
% Perform loop twice to prevent first item without domain
zerodomaincount = 0;
for i = 1:1:length(dir)
    dircheck = mat2str(dir(i));
    disp(dircheck)
    if dircheck(3:4) == ':\'
        domain = dircheck(2:4); % Pickup domain here
    else
        zerodomaincount = zerodomaincount+1;
    end
end
disp(append('Found ',num2str(zerodomaincount),' directory(s) without disk domain, automatically add domain if needed'))

for i = 1:1:length(dir)
    dircheck = mat2str(dir(i));
    if dircheck(3:4) == ':\'
        pause(0) % Bypass here
    else
        dir(i) = append(domain,dir(i));
    end
end
disp('Done')

disp(append('Total of ',num2str(length(dir)),' files to be exchanged: '))
disp(append(file,' -----> ',dir))

% Second directory
dirtwo = append(source,extractAfter(dir,2)); % Getting only directory without disk domain

% File availability check
for i = 1:1:length(file)
    if isfile(file(i))
        pause(0)
    else
        error(append('[!} ',file(i),' cannot be found in directory!'))
    end
end

while(1)
    choice = input('Do you want to proceed? [Y/N] ','s');
    if choice == "y"||choice == "Y"||choice == "yes"||choice == "Yes"||choice == "YES"
        disp(append('Copying ',num2str(length(file)),' files to folder...'))
        
        for i = 1:1:length(file) % Copy files to designated/specified location
            if not(isfolder(dir(i)))
                mkdir(dir(i))
                disp(append('Folder has been created at ',dir(i)))
            end
            copyfile(file(i),dir(i),'f')
            disp(append(file(i),' has been copied to ',dir(i)))
        end
        
        disp('Done')
        
        %{
        %This part writes status to table and save
        disp("Changing file status to 'Processing'")
        tab2 = readtable(excelfile);
        for i = 1:1:height(tab2)
            if char(table2cell(tab2(i,7))) == 'open' % Convers table cell to string
                tab2(i,7) = {'processing'};
                tab2(i,6) = {0};
            else
            end
        end
        writetable(tab2,'FileExchange.xlsx')
        disp(append('Status change saved to ',defaulttable))
        %}
        
        while(1)
            disp('The following procedure creates a folder structure under current directory and copied file into it.')
            choice2 = input('Do you want to proceed? [Y/N] ','s');
            if choice2 == "y"||choice2 == "Y"||choice2 == "yes"||choice2 == "Yes"||choice2 == "YES"
                % Following line creates folder structure (all folder and
                % subfolder)
                foldersource = dircheck;
                system(append('robocopy ',append(foldersource(1:10),'" "'),defaultsource,foldersource(4:10),'" /e /xf *'));
                
                folderdir = strings(length(dir),1);
                for i = 1:1:length(dir)
                    temp = mat2str(dir(i));
                    folderdir(i) = append(defaultsource,temp(4:end-1));
                end
                
                for i = 1:1:length(file) % Copy files to new folder structure
                    copyfile(file(i),folderdir(i),'f')
                end
                
                disp('Deleting Empty Folder...')
                %The following robocopy command deletes all empty folder
                system(append('robocopy "',defaultsource,foldersource(4:10),'" "',defaultsource,foldersource(4:10),'" /S /move'));
                
                for i = 1:1:length(file) % This loop shows copy status
                    disp(append(file(i),' has been copied to  ',folderdir(i)))
                end
                disp(append('[!] Folder structure copy has been created at ',defaultsource))
                break
            elseif choice2 == "n"||choice2 == "N"||choice2 == "no"||choice2 == "No"||choice2 == "NO"
                disp('All files has been copied to place, but folder structure and copy not created.')
                break
            else
                disp('Please input Y/N') % Keep asking if input is not yes or no
            end
        end
        
        disp('Program ended.')
        break
        
    elseif choice == "n"||choice == "N"||choice == "no"||choice == "No"||choice == "NO"
        disp('Program terminated.')
        break
        
    else
        disp('Please input Y/N') % Keep asking if input is not yes or no
    end
end

