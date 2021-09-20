clc; clear;
% nc_Files = 'g4.areaAvgTimeSeries.AIRX3STM_006_RelHumid_A.1000hPa.20020901-20160930.2E_2N_5E_7N.nc'; % the '.nc' files in the current folder
nc_Files='g4.areaAvgTimeSeries.AIRX3STM_7_0_Temperature_A.1000hPa.20020901-20161031.2E_2N_5E_7N.nc';
fileName='NCFileData11.xlsx'; % the excel file name to export to

%% File information
NCfiles_info=ncinfo(nc_Files); % collecting information on the .nc file
File_Name={NCfiles_info.Filename};
File_Format ={NCfiles_info.Format};
File_Attribute = table({NCfiles_info.Attributes.Name;NCfiles_info.Attributes.Value}');
No_of_Variables=length(NCfiles_info.Variables);
File_info=rows2vars(table(File_Name,File_Format,No_of_Variables));
writetable(File_info,fileName, 'Sheet','File_information','writemode', 'overwritesheet'); % write the file info to Sheet 1)
writetable(File_Attribute,fileName, 'Sheet','Other_File_information','writemode', 'overwritesheet');

%% Variable information
% No_of_Variables=length(NCfiles_info.Variables);
Variable_list={NCfiles_info.Variables.Name}';
Variable_Size={NCfiles_info.Variables.Size}';
Variable_Datatype={NCfiles_info.Variables.Datatype}';
Variable_FillValue={NCfiles_info.Variables.FillValue}';
Variable_info=table(Variable_list,Variable_Size,Variable_Datatype,Variable_FillValue);
writetable(Variable_info,fileName, 'Sheet','Variables_information','writemode','overwritesheet');

% writecell({'File Name', 'File Format'; NCfiles_info.Filename,NCfiles_info.Format}','Vatt.xlsx')

for i=1:length(NCfiles_info.Variables)
    Variable(i).sheetName=NCfiles_info.Variables(i).Name;
    Variable(i).Attribute=table({NCfiles_info.Variables(i).Attributes.Name;NCfiles_info.Variables(i).Attributes.Value}');
    writetable(Variable(i).Attribute,fileName, 'Sheet',[Variable(i).sheetName,' Atrb'],'writemode','overwritesheet');
    Variable(i).Data=table(ncread(nc_Files,NCfiles_info.Variables(i).Name));
    writetable(Variable(i).Data,fileName, 'Sheet',[Variable(i).sheetName,' Data'],'writemode','overwritesheet');
end

Year_N_month=string(table2array(Variable(3).Data)); %from inner loop, convert table to array, then to string
Year=str2num(cell2mat(cellfun(@(Year_N_month) Year_N_month(1:4), Year_N_month, 'Uniform',0))); %from inner brcket, extract year(4 xter) from the value which give cell array; then conver to Matrix then to number
Month=str2num(cell2mat(cellfun(@(Year_N_month) Year_N_month(5:end), Year_N_month, 'Uniform',0)));
RequireData=zeros(range(Year)+1,12); %predefining require date by numbers of years and 12 month in each year

NeededData=table2array(Variable(1).Data);
Year_Month_NeedData=[Year, Month, NeededData];
Year_Index=min(Year):max(Year);

for i=1:length(Year_Index)
    for j=1:12 % 12 months in a year
        for k=1:length(Year_Month_NeedData)
            if Year_Month_NeedData(k)==Year_Index(i) && Year_Month_NeedData(k,2)==j % This part get the year and the month
                RequireData(i,j)=Year_Month_NeedData(k,3); % This collect the needed data located in collum 3 after getting the year and the month j
            end
        end
    end
end