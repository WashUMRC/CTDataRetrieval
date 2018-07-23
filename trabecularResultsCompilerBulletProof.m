% function trabecularResultsCompilerBulletProof()
clear all;clc;
keyword = '*3D*MORPHO*.TXT*';%Identify which text files to grab

storeDir = uigetdir(pwd,'Please select a directory in which to store your text files');%set a folder to use as a target for mget
cd(storeDir);
sysLine = 'md scratch';
system(sysLine);
cd('scratch');

answer = inputdlg('Please enter 1 for micro or 2 for viva');%choose machine
reply = str2num(answer{1});

h = msgbox('Connecting');
% connect to machine and get to correct data directory
if reply == 1
    f = ftp('10.21.24.204','microct','mousebone4');
    ascii(f);
    disp(f)
    cDir = cd(f,'dk0');
    cDir = cd(f,'data');
elseif reply == 2
    f = ftp('10.21.24.203','microct','mousebone4');
    ascii(f);
    disp(f)
    cDir = cd(f,'dk0');
    cDir = cd(f,'data');
end
clear answer reply
delete(h);

%identify sample of interest
answer = inputdlg('Please enter your sample number');
tsto = '00000000';
len = 8 - length(answer{1});
sample = strcat(tsto(1:len),answer{1});

%identify measurements of interest
answer = inputdlg('Would you like to compile results for all measurements? Enter 1 for yes or 2 for no');
reply = str2num(answer{1});

h = msgbox('Getting files');
if reply == 1
    cDir = cd(f,sample);
    directories = dir(f);
    for i = 1:length(directories)
        tf = directories(i).isdir;
        if tf == 1
            cSubDir = cd(f,directories(i).name(1:length(directories(i).name)-6));
            file1 = dir(f,strcat(cSubDir,keyword));
            if ~isempty(file1)
                mget(f,file1.name);
            end
        else
        end
        cd(f,'..');
    end
elseif reply == 2
    cDir = cd(f,sample);
    answer = inputdlg('Please enter a list of measurements you would like to compile results for separated by commas with no spaces');
    measurements = str2num(answer{1});
    for i = 1:length(measurements)
        chars = length(num2str(measurements(i)));
        zs = zeros(1,8-chars);
        z=char();
        for j = 1:length(zs)
            z = strcat(z,num2str(zs(j)));
        end
        subDir = strcat(z,num2str(measurements(i)));
        cSubDir = cd(f,subDir);
        file1 = dir(f,strcat(cSubDir,keyword));
        mget(f,file1.name);
        cd(f,'..');
    end
end
delete(h);

%create header for output file
%create header for output file
fullHeader = {...
    'SampName',...
    'SampNo',...
    'MeasNo',...
    'MeasDate',...
    'ListDate',...
    'Filename',...
    'S-DOB',...
    'S-Remark',...
    'Meas-Rmk',...
    'Site',...
    'Energy-I-Code',...
    'Integr.Time',...
    'ControlfileNo',...
    'Ctrlf-Name',...
    'Sigma',...
    'Support',...
    'Threshold',...
    'Unit',...
    'Data-Threshold',...
    'VOX-TV',...
    'VOX-BV',...
    'VOX-BV/TV',...
    'Connectivity Density (Conn-Dens.)',...
    'Structure Model Index (TRI-SMI)',...
    'Trabecular Number (DT-Tb.N)',...
    'Trabecular Thickness (DT-Tb.Th)',...
    'Trabecular Separation (DT-Tb.Sp)',...
    'DT-Tb.(1/N).SD',...
    'Local Standard Deviation of Trabecular Thickness (DT-Tb.Th.SD)',...
    'Local Standard Deviation of Trabecular Spacing (DT-Tb.Sp.SD)',...
    'Bone Mineral Density (Apparent Density;mean density of total analyzed volume)',...
    'Tissue Mineral Density (density of thresholded tissue/thresholded volume)',...
    'Mean3',...
    'Mean4',...
    'Mean5',...
    'Mean-Units',...
    'TRI-TV',...
    'TRI-BV',...
    'TRI-BV/TV',...
    'TRI-BS',...
    'TRI-BS/BV',...
    'TRI-Tb.N',...
    'TRI-Tb.Th',...
    'TRI-Tb.Sp',...
    'TRI-DA',...
    'TRI-|H1|',...
    'TRI-|H2|',...
    'TRI-|H3|',...
    'TRI-H1x',...
    'TRI-H1y',...
    'TRI-H1z',...
    'TRI-H2x',...
    'TRI-H2y',...
    'TRI-H2z',...
    'TRI-H3x',...
    'TRI-H3y',...
    'TRI-H3z',...
    'Voxel Size X',...
    'Voxel Size Y',...
    'Voxel Size Z',...
    'Dim-X',...
    'Dim-Y',...
    'Dim-Z',...
    'Pos-X',...
    'Pos-Y',...
    'Pos-Z',...
    'MeasNumSlices',...
    'OperatorMeas',...
    'OperatorEval',...
    'CTDI[mGy]',...
    'RAW-Dir',...
    'RAW-Label',...
    'IMA-Dir',...
    'IMA-Label',...
    'ScannerID'...
    };

h = msgbox('Working');
txtList = dir([pwd '\' keyword]);

useVector = [1 2 3 4 5 6 8 9 12 14 63 66 15 16 17 20 21 22 23 24 25 26 27 29 30 31 32 36 60];

%Read in all the relevant data from the 3D analysis

sysLine = ['del "' pwd '\*.xls*'];
system(sysLine);

excel = actxserver('Excel.Application');
set(excel,'Visible',0);
for i = 1:length(txtList)
    workbook = excel.Workbooks;
    invoke(workbook,'Open',[pwd '\' txtList(i).name]);
    excel.ActiveWorkbook.SaveAs([pwd '\' txtList(i).name(1:length(txtList(i).name)-5) 'xlsx'],51);
end
invoke(excel, 'Quit');
delete(excel);

xlsList = dir([pwd '\' keyword(1:length(keyword)-8) '*.xlsx']);
%grab meaningful data from excel files
c=0;
for i = 1:length(xlsList)
    d=0;
    clear newRaw raw num
    [num txt raw] = xlsread([pwd '\' xlsList(i).name]);
    [a b] = size(raw);
    for j = 2:a
        if strcmpi(raw{j,1},'Calibration') == 0
%             e=0;
%             for k = 1:b
%                 if ~isempty(find(useVector == k))
%                     e=e+1;
%                     newRaw{j,e} = raw{j,k};
%                 end
%             end
            newRaw(j,:) = raw(j,useVector);
        end
    end
        [a b] = size(newRaw);
        for j = 2:a
            c=c+1;
            d=d+1;
            for k = 1:b
                data{c,k} = newRaw{j,k};
                if isnumeric(data{c,k})
                    data1{c,k} = num2str(data{c,k});
                    
                else
                    data1{c,k} = data{c,k};
                    
                end
            end
        end
    end
    numLines(i) = d;
% end
for i = 1:length(useVector)
    header{i} = fullHeader{1,useVector(i)};
end

header = fullHeader(useVector);

headOut = header;%(indexForOutput);
datOut = data1;%(:,indexForOutput);
[a b] = size(datOut);
delete(h);

h = msgbox('Writing data');
%write out data
fid = fopen([storeDir '\' sample ' Cancellous Results.txt'],'w');
for i = 1:length(headOut)
    if i < length(headOut)
        fprintf(fid,'%s\t',headOut{i});
    else
        fprintf(fid,'%s\n',headOut{i});
    end
end

for i = 1:a
    for j = 1:b
        %make numbers strings to print
        out = datOut{i,j};
        if ischar(out) ~= 1
            out = num2str(out);
        end
        if j < b
            fprintf(fid,'%s\t',out);
        else
            fprintf(fid,'%s\n',out);
        end
    end
end

fclose(fid);

delete(h);