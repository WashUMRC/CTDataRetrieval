function CollectOutputsFromContouringGUI()

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%This script compiles data from output text files derived from
%%contouringGUI.m  It expects that you've kept your data in the same
%%file structure as is generated when you retrieve DICOMs with the suite's
%%tools in DICOMManagement.
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
pathstr = uigetdir(pwd,'Please select the parent folder containing your measurement folders.');

%first collect list of text file names to be aggregated.
dirs = dir(pathstr);
for i = 3:length(dirs)
    if dirs(i).isdir == 1
        thisDir = fullfile(pathstr,dirs(i).name);
        txtFiles = dir(fullfile(thisDir,'*.txt'));
        if length(txtFiles) > 0
            for j = 1:length(txtFiles)
                [filepath{j},name{j},ext{j}] = fileparts(fullfile(thisDir,txtFiles(j).name));
            end
            names = unique(name);
        end
    end
end
if length(names) > 0
    names = unique(names);
end
        
%Iterate through by text file name, looking in each measurement folder and
%aggregating data
for i = 1:length(names)
    ct=0;
    dirs = dir(pathstr);
    for j = 3:length(dirs)
        if dirs(j).isdir == 1
            thisDir = fullfile(pathstr,dirs(j).name);
            thisTxt = fullfile(thisDir,[names{i} '.txt']);
            thisTxts = dir(thisTxt);
            if length(thisTxts) > 0
                ct=ct+1;
                excel = actxserver('Excel.Application');
                set(excel,'Visible',0);
                workbook = excel.Workbooks;
                invoke(workbook,'Open',thisTxt);
                excel.ActiveWorkbook.SaveAs([thisTxt(1:end-3) '.xlsx'],51);
                invoke(excel, 'Quit');
                delete(excel);
                [~,~,raw] = xlsread([thisTxt(1:end-3) '.xlsx']);
                delete([thisTxt(1:end-3) '.xlsx.']);
                if ct == 1
                    fid = fopen(fullfile(pathstr,[names{i} 'Compiled.txt']),'w');
                    [a,b] = size(raw);
                    for k = 1:a
                        for l = 1:b
                            toPrint = cell2mat(raw(k,l));
                            if l ~= b
                                try
                                    fprintf(fid,'%s\t',num2str(toPrint));
                                catch
                                    fprintf(fid,'%s\t',toPrint);
                                end
                            else
                                try
                                    fprintf(fid,'%s\n',num2str(toPrint));
                                catch
                                    fprintf(fid,'%s\n',toPrint);
                                end
                            end
                        end
                    end
                    fclose(fid);
                else
                     fid = fopen(fullfile(pathstr,[names{i} 'Compiled.txt']),'a');
                    [a,b] = size(raw);
                    for k = 2:a
                        for l = 1:b
                            toPrint = cell2mat(raw(k,l));
                            if l ~= b
                                try
                                    fprintf(fid,'%s\t',num2str(toPrint));
                                catch
                                    fprintf(fid,'%s\t',toPrint);
                                end
                            else
                                try
                                    fprintf(fid,'%s\n',num2str(toPrint));
                                catch
                                    fprintf(fid,'%s\n',toPrint);
                                end
                            end
                        end
                    end
                    fclose(fid);
                end
            end
        end
    end
end
            