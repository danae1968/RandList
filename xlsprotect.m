function xlsprotect(file,option,varargin)

%XLSPROTECT Protects/Unprotects Selected Sheet/Workbook with protection options.
%
% SYNTAX
%
% FILE PROTECTION:
% xlsprotect(file,'protect_file',pwd2open,pwd2modify,read_only_recommended,create_backup)
%
% xlsprotect(file,'unprotect_file',pwd2open)
% xlsprotect(file,'unprotect_file',pwd2open,pwd2modify)
%
%
%           DISCRIPTION
%                   file:       Name of Excel File.
%                   pwd2open:   Password consisting of string characters
%                               used to open file.
%                   pwd2modify: Password consisting of string characters
%                               used to modify file.
%                   read_only_recommended:  (1) Yes. (0) No.
%                   create_backup:  (1) Yes. (0) No.
%
% SHEET PROTECTION:
% xlsprotect(file,'protect_sheet',sheetname,password)
% xlsprotect(file,'protect_sheet',sheetname,password,s_options)
% xlsprotect(file,'protect_sheet',sheetname,s_options)
% xlsprotect(file,'protect_sheet',sheetname)
%
% xlsprotect(file,'unprotect_sheet',sheetname)
% xlsprotect(file,'unprotect_sheet',sheetname,password)
%
%           DISCRIPTION
%                   file:           Name of Excel File.
%                   sheetname:      Name os Sheet to be protected/unprotected.
%                   password:       Password consisting of string characters.
%                   s_options:      Sheet Protection Options, row of zeros and ones
%                       (ex. [1 0 1 1 0 ...]) with maximum of 16 switches.
%                       Switches are in the following order:
%                       1- [DrawingObjects] (0/1)
%                       2- [Contents] (0/1)
%                       3- [Scenarios] (0/1)
%                       4- [UserInterfaceOnly] (0/1)
%                       5- [AllowFormattingCells] (0/1)
%                       6- [AllowFormattingColumns] (0/1)
%                       7- [AllowFormattingRows] (0/1)
%                       8- [AllowInsertingColumns] (0/1)
%                       9- [AllowInsertingRows] (0/1)
%                       10-[AllowInsertingHyperlinks] (0/1) 
%                       11-[AllowDeletingColumns] (0/1)
%                       12-[AllowDeletingRows] (0/1)
%                       13-[AllowSorting] (0/1)
%                       14-[AllowFiltering] (0/1)
%                       15-[AllowUsingPivotTables] (0/1)
%
% WORKBOOK PROTECTION:
% xlsprotect(file,'protect_workbook',password)
% xlsprotect(file,'protect_workbook')
%
% xlsprotect(file,'unprotect_workbook')
% xlsprotect(file,'unprotect_workbook',password)
%
%           DISCRIPTION
%
%                   file:       Name of Excel File.
%                   password:   Password consisting of string characters.
%
% Examples:
% 
%      xlsprotect('data.xls','protect_file','ThisIsMyPassword','',0,0)
%      xlsprotect('data.xls','unprotect_file','ThisIsMyPassword','')
%
%      xlsprotect('data.xls','protect_sheet','Sheet1');
%      xlsprotect('data.xls','protect_sheet','Sheet1','ThisIsMyPassword');
%      xlsprotect('data.xls','protect_sheet','Sheet1','ThisIsMyPassword');
%      xlsprotect('data.xls','protect_sheet','Sheet1','ThisIsMyPassword',[0 1 0 0 1 1]);
%      xlsprotect('data.xls','unprotect_sheet','Sheet1','ThisIsMyPassword');
%
%      xlsprotect('data.xls','protect_workbook');
%      xlsprotect('data.xls','protect_workbook','ThisIsMyPassword');
%      xlsprotect('data.xls','unprotect_workbook','ThisIsMyPassword');
%

%   Copyright 2004 Fahad Al Mahmood
%   Version: 1.0 $  $Date: 12-Oct-2004
%   Version: 1.5 $  $Date: 28-Nov-2004  (File Protection Added)
%   Modified by Danae Papadopetraki 2019


% Setting up the file name with path

[fpath,fname,fext] = fileparts(file);
if isempty(fpath)
    out_path = pwd;
elseif fpath(1)=='.'
    out_path = [pwd filesep fpath];
else
    out_path = fpath;
end
file = [out_path filesep fname fext];
        
% Opening Excel File
Excel = actxserver('Excel.Application');
set(Excel,'Visible',0);

% ---------------
% Protect File
% ---------------
if strcmp(option,'protect_file')
    Workbook = invoke(Excel.Workbooks,'open',file);
    invoke(Excel.ActiveWorkbook,'SaveAs',[out_path filesep 'temp000.xls'],[],varargin{1},varargin{2},varargin{3},varargin{4},1,1,1,1,1);
    invoke(Excel,'Quit');
    delete(Excel);
    eval(['delete ''' file '''']);
     movefile([out_path filesep 'temp000.xls'], file,'f');    
% ---------------
% Unprotect File
% ---------------    
elseif strcmp(option,'unprotect_file')
    if length(varargin)==1
        Workbook = invoke(Excel.Workbooks,'open',file,0,0,1,varargin{1});
    elseif length(varargin)==2
        Workbook = invoke(Excel.Workbooks,'open',file,0,0,1,varargin{1},varargin{2});
    end
    invoke(Excel.ActiveWorkbook,'SaveAs',[out_path filesep 'temp000.xls'],[],'','',0,0,1,1,1,1,1);
    invoke(Excel,'Quit');
    delete(Excel);
    eval(['delete ''' file '''']);
    movefile([out_path filesep 'temp000.xls'], file,'f');
    
    
% ---------------
% Protect Sheet
% ---------------
elseif strcmp(option,'protect_sheet')
    Workbook = invoke(Excel.Workbooks,'open',file);
    sheetname = varargin{1};
    Sheets = Excel.ActiveWorkBook.Sheets;
    sheet = get(Sheets,'Item',sheetname);
    invoke(sheet, 'Activate');
    op = zeros(1,16);
    op(2) = 1;
    op(4) = 1;
    op(15) = 1;
    op(16) = 1;
    if length(varargin)==3
        password = varargin{2};
        s_options = varargin{3};
        for k=1:length(s_options)
            op(k) = s_options(k);
        end
    elseif length(varargin)==2
        if isnumeric(varargin{2})
            s_options = varargin{2};
            password = '';
            for k=1:length(s_options)
                op(k) = s_options(k);
            end
        else
            password = varargin{2};
        end
    elseif length(varargin)==1
        password = '';
    end
    invoke(Excel.ActiveSheet,'protect',password,...
        op(1),op(2),op(3),op(4),op(5),op(6),...
        op(7),op(8),op(9),op(10),op(11),...
        op(12),op(13),op(14),op(15));
    invoke(Workbook, 'Save');
    invoke(Excel,'Quit');
    delete(Excel);
    

% ---------------
% Unprotect Sheet
% ---------------
elseif strcmp(option,'unprotect_sheet')
    Workbook = invoke(Excel.Workbooks,'open',file);
    sheetname = varargin{1};
    Sheets = Excel.ActiveWorkBook.Sheets;
    sheet = get(Sheets,'Item',sheetname);
    invoke(sheet, 'Activate');
    if length(varargin)==2
        password = varargin{2};
    else
        password = '';
    end
    
    try
        invoke(Excel.ActiveSheet,'unprotect',password);
    catch
        invoke(Workbook, 'Save');
        invoke(Excel,'Quit');
        delete(Excel);
        error('The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.');
    end
    invoke(Workbook, 'Save');
    invoke(Excel,'Quit');
    delete(Excel);
    
% ---------------
% Protect Workbook
% ---------------

elseif strcmp(option,'protect_workbook')
    Workbook = invoke(Excel.Workbooks,'open',file);
    if ~isempty(varargin)
        password = varargin{1};
    else
        password = '';
    end
    invoke(Excel.ActiveWorkbook,'protect',password);
    invoke(Workbook, 'Save');
    invoke(Excel,'Quit');
    delete(Excel);

% -----------------
% Unprotect Workbook
% -----------------  

elseif strcmp(option,'unprotect_workbook')
    Workbook = invoke(Excel.Workbooks,'open',file);
    if ~isempty(varargin)
        password = varargin{1};
    else
        password = '';
    end
    try
        invoke(Excel.ActiveWorkbook,'unprotect',password);
    catch
        invoke(Workbook, 'Save');
        invoke(Excel,'Quit');
        delete(Excel);
        error('The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.');
    end
    invoke(Workbook, 'Save');
    invoke(Excel,'Quit');
    delete(Excel);
end