function corticalResultsCompilerScript36()

keyword1 = '*3DRESULTS_CORT_W_THICK.TXT*';%Identify which text files to grab
keyword2 = '*MOIRESULTS_CORT_W_THICK.TXT*';

storeDir = uigetdir(cd,'Please select a directory in which to store your text files');%set a folder to use as a target for mget
cd(storeDir);
sysLine = 'md scratch';
system(sysLine);
cd('scratch');

answer = inputdlg('Please enter 1 for micro or 2 for viva');%choose machine
reply = str2num(answer{1});

% connect to machine and get to correct data directory
if reply == 1
    f = ftp('10.21.24.204','microct','mousebone4','System','OpenVMS');
    ascii(f);
    disp(f)
    cDir = cd(f,'dk0');
    cDir = cd(f,'data');
elseif reply == 2
    f = ftp('10.21.24.203','microct','mousebone4','System','OpenVMS');
    ascii(f);
    disp(f)
    cDir = cd(f,'dk0');
    cDir = cd(f,'data');
end
clear answer reply

%identify sample of interest
answer = inputdlg('Please enter your sample number');
tsto = '00000000';
len = 8 - length(answer{1});
sample = strcat(tsto(1:len),answer{1});

%identify measurements of interest
answer = inputdlg('Would you like to compile results for all measurements? Enter 1 for yes or 2 for no');
reply = str2num(answer{1});

if reply == 1
    cDir = cd(f,sample);
    directories = dir(f);
    for i = 1:length(directories)
        tf = directories(i).isdir;
        if tf == 1
            cSubDir = cd(f,directories(i).name(1:length(directories(i).name)-6));
            file1 = dir(f,strcat(cSubDir,keyword1));
            file2 = dir(f,strcat(cSubDir,keyword2));
            if ~isempty(file1) && ~isempty(file2)
                mget(f,file1.name);
                mget(f,file2.name);
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
        file1 = dir(f,strcat(cSubDir,keyword1));
        file2 = dir(f,strcat(cSubDir,keyword2));
        mget(f,file1.name);
        mget(f,file2.name);
        cd(f,'..');
    end
end

%get rid of ;N
fileRenamerSub(pwd);

%create header for output file
Results_MIDSH_fullHeader = {...
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
    'Conn-Dens.',...
    'TRI-SMI',...
    'DT-Ct.N',...
    'DT-Ct.Th',...
    'DT-Ct.Sp',...
    'DT-Ct.(1/N).SD',...
    'DT-Ct.Th.SD',...
    'DT-Ct.Sp.SD',...
    'vBMD',...
    'TMD',...
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
    'El-Siz-X',...
    'El-Siz-Y',...
    'El-Siz-Z',...
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
useVector1 = [1 2 3 8 9 12 14 15 16 17 20 21 22 23 24 26 29 31 32 36 58 59 60];
slices = 63;
MOI_fullHeader = {...
    'Patient-Name',...
    'S-No',...
    'M-No',...
    'ListDate',...
    'Segmentation',...
    'El_size_mm',...
    'CMx[mm]',...
    'CMy[mm]',...
    'Ixx[mm^4]',...
    'Iyy[mm^4]',...
    'Ixy[mm^4]',...
    'pMOI[mm^4]',...
    'Ixx/Cy[mm^3]',...
    'Iyy/Cx[mm^3]',...
    'Imax[mm^4]',...
    'Imin[mm^4]',...
    'Angle[deg]',...
    'Imax/Cmax[mm^3]',...
    'Imin/Cmin[mm^3]',...
    'BArea[mm^2]',...
    'TArea[mm^2]',...
    'BA/TA[1]',...
    'TRI-Ct.Th',...
    'Mean1',...
    'Mean1SD',...
    'Mean2',...
    'Mean2SD',...
    'DT-Ct.Th+',...
    'DT-Ct.Th.SD',...
    'DT-Ct.Sp+',...
    'DT-Ct.Sp.SD'...
    };
useVector2 = [12]; 

txtList1 = dir([pwd '\' keyword1]);
txtList2 = dir([pwd '\' keyword2]);

%use excel to parse text files
excel = actxserver('Excel.Application');
set(excel,'Visible',0);
for i = 1:length(txtList1)
    workbook = excel.Workbooks;
    invoke(workbook,'Open',[pwd '\' txtList1(i).name]);
    excel.ActiveWorkbook.SaveAs([pwd '\' txtList1(i).name(1:length(txtList1(i).name)-3) 'xlsx'],51);
end
invoke(excel, 'Quit');
delete(excel);

%use excel to parse text files
excel = actxserver('Excel.Application');
set(excel,'Visible',0);
for i = 1:length(txtList2)
    workbook = excel.Workbooks;
    invoke(workbook,'Open',[pwd '\' txtList2(i).name]);
    excel.ActiveWorkbook.SaveAs([pwd '\' txtList2(i).name(1:length(txtList2(i).name)-3) 'xlsx'],51);
end
invoke(excel, 'Quit');
delete(excel);

xlsList1 = dir([pwd '\*' keyword1(1:length(keyword1)-8) '*.xlsx']);
xlsList2 = dir([pwd '\*' keyword2(1:length(keyword2)-8) '*.xlsx']);

%grab meaningful data from excel files
c=0;
for i = 1:length(xlsList1)
    d=0;
    clear newRaw raw num
    [num txt raw] = xlsread([pwd '\' xlsList1(i).name]);
    [a b] = size(raw);
   
    for j = 1:a
        e=0;
        for k = 1:b
            if ~isempty(find(useVector1 == k))
                e=e+1;
                newRaw{j,e} = raw{j,k};
            end
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
        zSlices(c) = raw{j,slices};
        zSlices(c) = raw{j,slices};
        TA(c) = str2num(data1{c,11}) / (zSlices(c)*str2num(data1{c,23})*1000);
        BA(c) = str2num(data1{c,12}) / (zSlices(c)*str2num(data1{c,23})*1000);
        MA(c) = TA(c) - BA(c);

    end
    numLines(i) = d;
end
for i = 1:length(useVector1)
    header1{i} = Results_MIDSH_fullHeader{1,useVector1(i)};
end
clear data newRaw num txt raw
c=0;
for i = 1:length(xlsList2)
    d=0;
    [num txt raw] = xlsread([pwd '\' xlsList2(i).name]);
    [a b] = size(raw);
   
    for j = 1:a
        e=0;
        for k = 1:b
            if ~isempty(find(useVector2 == k))
                e=e+1;
                newRaw{j,e} = raw{j,k};
            end
        end
    end
    [a b] = size(newRaw);
    for j = 2:a
        c=c+1;
        d=d+1;
        for k = 1:b
            data{c,k} = newRaw{j,k};
            if isnumeric(data{c,k})
                data2{c,k} = num2str(data{c,k});
            else
                data2{c,k} = data{c,k};
            end
        end
    end
end
for i = 1:length(useVector2)
    header2{i} = MOI_fullHeader{1,useVector2(i)};
end

c=0;
for i = 1:length(header1)
    c=c+1;
    header{c} = header1{i};
end
twoDHeader = {'Total Area','Bone Area','Medullary Area'};
for i = 1:3
    c=c+1;
    header{c} = twoDHeader{i};
end
for i = 1:length(header2)
    c=c+1;
    header{c} = header2{i};
end

c=0;
[a b] = size(data1);
for i = 1:a
    c=0;
    for j = 1:b
        c=c+1;
        data{i,c} = data1{i,j};
    end
    c=c+1;
    data{i,c} = num2str(TA(i));
    c=c+1;
    data{i,c} = num2str(BA(i));
    c=c+1;
    data{i,c} = num2str(MA(i));
end

[a, b] = size(data2);
 c=c+1;
for i = 1:a
%     c=0;
    for j = 1:b
%         c=c+1;
        data{i,c} = data2{i,j};
    end
end

%Time to print
fid = fopen([storeDir '\' sample ' Cortical Results.txt'],'w');
for i = 1:length(header)
    if i < length(header)
        fprintf(fid,'%s\t',header{i});
    else
        fprintf(fid,'%s\n',header{i});
    end
end
c=0;
for i = 1:length(numLines)
    for j = 1:numLines(i)
        c=c+1;
        a = length(data(c,:));
        for k = 1:a
            if ~isnan(data{c,k})
                fprintf(fid,'%s\t',data{c,k});
            else
                fprintf(fid,'%s\t','');
            end
        end
        fprintf(fid,'%s\n','');
    end
    fprintf(fid,'%s\n','');
end
fclose(fid);
msgbox('Data successfully coalated! There should be a new text file in the directory you chose earlier.');
% pause(2);
fclose('all');

function fileRenamerSub(directory)
    os = computer;
    if ~isempty(strfind(os,'PC'))
        ind=1;
        files = dir(strcat(directory,'\*;*'));
        for i = 1:length(files)
            sysline = char(strcat('ren',{' '},'"',directory,'\',files(i).name,'"',{' '},files(i).name(1:length(files(i).name)-2)));
            system(sysline);
            ind=2;
        end
        if ind ~= 2
            files = dir(strcat(directory,'\*_*'));
            for i = 1:length(files)
                movefile([directory '\' files(i).name], [directory '\' files(i).name(1:length(files(i).name)-2)]);
            end
        end
    elseif ~isempty(strfind(os,'MAC'))
        ind=1;
        file = dir(strcat(directory,'/*_*'));
        for i = 1:length(files)
            movefile([directory '/' file(i).name], [directory '\' file(i).name(1:length(file(i).name)-2)]);
            ind=2;
        end
        if ind ~= 2
            file = dir(strcat(directory,'/*;*'));
            for i = 1:length(files)
                movefile([directory '/' file(i).name], [directory '\' file(i).name(1:length(file(i).name)-2)]);
            end
        end
    end
    msgbox('Files renamed!');
    pause(0.5);
    close('all');
