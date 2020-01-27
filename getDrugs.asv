%if original name of script save a copy, if final not
clear 

if strcmp(mfilename,'getDrugs')
saveScript=1;
elseif strcmp(mfilename,'getDrugsF')
    saveScript=0;
end

% get date and time
t=now;
dateNow = datetime(t,'ConvertFrom','datenum');

% windows or linux
if ispc
    projectDir = '\\fileserver.dccn.nl\project\3024005.02\';
elseif isunix
    projectDir = '/project/3024005.02/';
else
    error('Unknown OS');
end

%% fill in here relevant pathways
codeDir=fullfile(projectDir,'TestingDay','MedicationPreparation');
fileDrugs=fullfile(codeDir,'randList.xlsx');
fileLogRun=fullfile(codeDir,'runLog.xlsx');

externalResearcherPath='C:\Users\danpap\Documents\GitHub\RandList';

%if this is the fist time the script is run, the runData file does not
%exist
if ~exist(fileLogRun, 'file')
    xlswrite(fileLogRun,{'date','user','subNo','session'},1,'A1:D1')
end
%% find users
if ispc
    username=getenv('username');
elseif isunix
    [~,username]=system('whoami');
    username(end)=[]; 
end

if strcmp(username,'evakli')
    userS = 'Eva';
elseif strcmp(username,'fellin')
    userS = 'Felix';
elseif strcmp(username,'danpap')
    userS = 'Danae';
elseif strcmp(username,'rosco')
    userS = 'Roshan';
elseif strcmp(username,'richel')
    userS = 'Rick';
elseif strcmp(username,'frenie')
    userS = 'Freek';
elseif strcmp(username,'anovhei')
    userS = 'Anouk';
    
else
    error('I am sorry %s you do not have access to use this code!',userS)
end

%% give information
fprintf('Hello %s! Another testing day? How exciting!\n',userS)
prompt = '\n Could you please tell me the subject number?\n';
subNo = input(prompt);

if ~ismember(subNo,1:30)
    error('Subject number must be between 1 and 30')
end

prompt2 = '\nCould you please tell me the session number?\n';
session = input(prompt2);

if ~ismember(session,1:2)
    error('Session number must be 1 or 2')
end


%% ADD EXCEL PASSWORD HERE
password='danae';

%% known password: do not modify this password!! Allows us to read, but not write the log file
known_password='pd';
%% activate activeX
Excel = actxserver('excel.application');

workbook = Excel.Workbooks.Open(fileDrugs, [], true, [], password);
resultSheet='Sheet1';

exlSheet1 = Excel.Sheets.Item(resultSheet);

robj = exlSheet1.Columns.End(4);       % Find the end of the column
numrows = robj.row;                    % And determine what row it is
dat_range = ['A1:C' num2str(numrows)]; % Read to the last row
rngObj = exlSheet1.Range(dat_range);

exlData = rngObj.Value;

codes=cell2mat(exlData);

%% close all
invoke(Excel,'Quit');
delete(Excel);

drugNum=codes(codes(:,1)==subNo,session+1);

switch drugNum
    case 1
        drugToday='PLACEBO';
    case 2
        drugToday='PROPRANOLOL';
end

fprintf('The drug for subject %d session %d is %s! Good luck with the experiment %s!',subNo,session,drugToday,userS)

runDataNow={char(dateNow) userS subNo session};
%% activate activeX for logfile
% Opening Excel File
xlsprotect(fileLogRun,'unprotect_file',known_password,password)

Excel = actxserver('Excel.Application');
set(Excel,'Visible',0);

 Workbook = invoke(Excel.Workbooks,'open',fileLogRun);

% Make the first sheet active
eSheets = Excel.ActiveWorkbook.Sheets;
eSheet1 = eSheets.get('Item', 1);
eSheet1.Activate
% find range
b=eSheet1.Columns.End(4);
numrows=b.row;

%if empty gives infinite number
if numrows>1000
    numrows=1;
end
dat_range = ['A1' ':' 'D' num2str(numrows)];

%read
rangeObj=eSheet1.Range(dat_range);
runData=rangeObj.Value;

%new data
runData=[runData;runDataNow];

%new range
dat_range_new = ['A1' ':' 'D' num2str(numrows+1)];
Range = eSheet1.get('Range',dat_range_new);

%save new data
Range.Value=runData;
%remove alerts
Excel.DisplayAlerts = 0;
invoke(Workbook, 'Save');

%close all
invoke(Excel,'Quit');
delete(Excel);    
    
%protect again and save
xlsprotect(fileLogRun,'protect_file',known_password,password,0,1)

%kill excel if bug
[taskstate, taskmsg] = system('tasklist|findstr "EXCEL.EXE"');
if ~isempty(taskmsg)
  status = system('taskkill /F /IM EXCEL.EXE');
end


%% if this is the first time the script is run, save a copy of this script with the password and protect the source code
if saveScript
    
scriptName=fullfile(codeDir,'getDrugsF.m');

if exist(scriptName,'file')
warning('Final version already exists')
end

copyfile(fullfile(codeDir,'getDrugs.m'),scriptName)
% no access to source code
pcode(scriptName)
end 

clear
