% function corticalResultsCompiler()

keyword1 = '*3DRESULTS_MIDSH_MORESLICES.TXT*';%Identify which text files to grab
keyword2 = '*MOIRESULTS_MIDSH_MORESLICES.TXT*';

storeDir = uigetdir('c:\users\microct\documents','Please select a directory in which to store your text files');%set a folder to use as a target for mget
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
delete(h);

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
    'DT-Ct.Th (mm)',...
    'DT-Ct.Sp (mm)',...
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
useVector1 = [1 2 3 4 5 8 9 12 14 15 16 17 20 21 22 23 24 26 29 31 32 36 60];
slices = [63,66];
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
headVec2 = [12];



h = msgbox('Working');

txtList1 = dir([pwd '\' keyword1]);
txtList2 = dir([pwd '\' keyword2]);

sysLine = ['del "' pwd '\*.xls*"'];
system(sysLine);

excel = actxserver('Excel.Application');
set(excel,'Visible',0);
for i = 1:length(txtList1)
    workbook = excel.Workbooks;
    invoke(workbook,'Open',[pwd '\' txtList1(i).name]);
    excel.ActiveWorkbook.SaveAs([pwd '\' txtList1(i).name(1:length(txtList1(i).name)-5) 'xlsx'],51);
end
invoke(excel, 'Quit');
delete(excel);

excel = actxserver('Excel.Application');
set(excel,'Visible',0);
for i = 1:length(txtList2)
    workbook = excel.Workbooks;
    invoke(workbook,'Open',[pwd '\' txtList2(i).name]);
    excel.ActiveWorkbook.SaveAs([pwd '\' txtList2(i).name(1:length(txtList2(i).name)-5) 'xlsx'],51);
end
invoke(excel, 'Quit');
delete(excel);

xlsList1 = dir([pwd '\' keyword1(1:end-5) '.xls*']);
xlsList2 = dir([pwd '\' keyword2(1:end-5) '.xls*']);


if length(txtList1) ~= length(txtList2)
    msgbox('You do not have an equal number of 2D and 3D analyses!');
    error('You do not have an equal number of 2D and 3D analyses!');
end




%Read in all the relevant data from the 3D analysis
c=0;
dataOut = cell(0);
for i = 1:length(txtList1)  %%put in 2d numbers, get header right
    [~,~,raw1] = xlsread([pwd '\' xlsList1(i).name]);
    [~,~,raw2] = xlsread([pwd '\' xlsList2(i).name]);
%     data = importdata([pwd '\' txtList1(i).name]);
%     data2 = importdata([pwd '\' txtList2(i).name]);
%     tmp = struct();
%     tmp.data = data2;
%     data2 = tmp.data;
    data = struct();
    data2 = struct();
    data.textdata = raw1;
    data2.textdata = raw2;
    
    %clean up failed analyses
    if length(data.textdata(:,1)) ~= length(data2.textdata(:,1))
        for j = 2:length(data2.textdata(:,1))
            try
                for k = 7:length(data2.textdata(j,7:end))
                    if isstr(cell2mat(data2.textdata(j,k)))
                        strrep(data2.textdata(j,k),'~','');
                    end
                    try
                        nums = cell2mat(data2.textdata(j,7:end));
                    catch
                        data2.textdata(j,:) = [];
                    end
                    
                end
            catch
            end
        end
        flag = [];
        for j = 1:length(data.textdata(:,1))
            for k = 1:length(data.textdata(1,:))
                if ~isempty(strfind(data.textdata{j,k},'****'))
                    flag = [flag j];
                end
            end
        end
        data.textdata(flag,:) = [];
    end
        
    if i == 1
        TA = cell2mat(data.textdata(2:end,20)) ./ (cell2mat(data.textdata(2:end,60)) .* cell2mat(data.textdata(2:end,63)));
        BA = cell2mat(data.textdata(2:end,21)) ./ (cell2mat(data.textdata(2:end,60)) .* cell2mat(data.textdata(2:end,63)));
        MA = TA - BA;
        dataOut(1:length(data.textdata(2:end,1)),:) = [data.textdata(2:end,[useVector1,slices]),...%pull numbers of interest from data
            num2cell(TA),num2cell(BA),num2cell(MA),...%2D info
            num2cell(data2.textdata(2:end,useVector2))];%grab slices analyzed information
        c=c+length(dataOut(:,1));
    else
        TA = cell2mat(data.textdata(2:end,20)) ./ (cell2mat(data.textdata(2:end,60)) .* cell2mat(data.textdata(2,63)));
        BA = cell2mat(data.textdata(2:end,21)) ./ (cell2mat(data.textdata(2:end,60)) .* cell2mat(data.textdata(2,63)));
        MA = TA - BA;
        dataOut(c+1:c+length(data.textdata(2:end,1)),:) = [data.textdata(2:end,[useVector1,slices]),...
            num2cell(TA),num2cell(BA),num2cell(MA),...%2D info...
            num2cell(data2.textdata(2:end,useVector2))];
        c=c+length(data.textdata(2:end,1));
    end 

        
end

twoDHeader = {'Total Area (mm^2)','Bone Area (mm^2)','Medullary Area (mm^2)'};

header = [Results_MIDSH_fullHeader([useVector1,slices]),twoDHeader,MOI_fullHeader(headVec2)];

indexForOutput = [1,2,3,4,5,6,7,8,9,10,11,12,18,19,21,22,26,27,28,29,23,24,25];

headOut = header(indexForOutput);
datOut = dataOut(:,indexForOutput);
[a b] = size(datOut);
delete(h);

h = msgbox('Writing data');
%write out data
fid = fopen([storeDir '\' sample ' Cortical Results.txt'],'w');
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
        if iscell(datOut{i,j})
            datOut{i,j} = cell2mat(datOut{i,j});
        end
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