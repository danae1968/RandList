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
codeDir=fullfile(projectDir,'Taskcode','RandList');addpath(codeDir)
drugDir=fullfile(projectDir,'TestingDay','MedicationPreparation');addpath(drugDir)
fileDrugs=fullfile(drugDir,'randList.xlsx');
fileLogRun=fullfile(drugDir,'runLog.xlsx');
fileDropOut=fullfile(drugDir,'DropOuts.xlsx');

%if this is the fist time the script is run, the runData file does not
%exist
if exist(fileLogRun, 'file')==0
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
elseif strcmp(username,'anovdhei')
    userS = 'Anouk';
elseif strcmp(username,'ninvlie')
    userS = 'Nina';
elseif strcmp(username,'marjoh')
    userS = 'Martin';
elseif strcmp(username,'vicsue')
    userS = 'Victoria';
elseif strcmp(username,'jortic')
    userS = 'Jorryt';
elseif strcmp(username,'chrisa')
    userS = 'Christina';
else   
    error('I am sorry %s you do not have access to use this code!',userS)
end

%% give information
if saveScript==0
    fprintf('Hello %s! Another testing day? How exciting!\n',userS)
else %message for person making the list
    fprintf('Hello %s! Make sure you have copy pasted the password of the excel file in this document!\n',userS)
    
end
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
xlsprotect(fileDrugs,'unprotect_file',password,password)
Excel = actxserver('excel.application');
set(Excel,'Visible',0);

workbook = Excel.Workbooks.Open(fileDrugs, [], true, [], password);
resultSheet='Sheet1';

exlSheet1 = Excel.Sheets.Item(resultSheet);

robj = exlSheet1.Columns.End(4);       % Find the end of the column
numrows = robj.row;                    % And determine what row it is
dat_range = ['A1:C' num2str(numrows)]; % Read to the last row
rngObj = exlSheet1.Range(dat_range);

exlData = rngObj.Value;

codes=cell2mat(exlData);

%% see if there have been any dropouts and adjast future sessions if dropouts don't cancel out
dropN=xlsread(fileDropOut);

if numel(dropN)>1
    testedSubs=codes(1:subNo-1,:);
    testedTrue=testedSubs(~ismember(testedSubs(:,1),dropN),:); %remove dropouts
    nonTestedSubs=codes(subNo:end,:);
    actualList=[testedTrue; nonTestedSubs];
    order=sum(actualList(:,2)==1)-sum(actualList(:,2)==2); %measure order counterbalancing after dropouts
    toAdjust=floor(abs(order/2));
    %find how many are left
    if toAdjust>0
        if order<=-2
            surplusOrder=nonTestedSubs(nonTestedSubs(:,2)==2,:);
            %if we don't have enough subjects of the order to be adjusted
            if length(surplusOrder)<=toAdjust
                toAdjust=length(surplusOrder);
            end
            adjSubs=randsample(surplusOrder(:,1),toAdjust);%subjects to adjust
            codes(ismember(codes(:,1),adjSubs),3)=2;
            codes(ismember(codes(:,1),adjSubs),2)=1;
            
            
        elseif order>=2
            surplusOrder=nonTestedSubs(nonTestedSubs(:,2)==1,:);
            %if we don't have enough subjects of the order to be adjusted
            if length(surplusOrder)<=toAdjust
                toAdjust=length(surplusOrder);
            end
            adjSubs=randsample(surplusOrder(:,1),toAdjust);%subjects to adjust
            codes(ismember(codes(:,1),adjSubs),3)=1;
            codes(ismember(codes(:,1),adjSubs),2)=2;
        end
        
    rngObj.Value=codes;
    Excel.DisplayAlerts = 0;
    invoke(workbook, 'Save');
    
    %close all
    invoke(Excel,'Quit');
    delete(Excel);
    
    %protect again and save adjasted list
    xlsprotect(fileDrugs,'protect_file',password,password,0,1)
    
    %kill excel if bug
    [taskstate, taskmsg] = system('tasklist|findstr "EXCEL.EXE"');
    if ~isempty(taskmsg)
        status = system('taskkill /F /IM EXCEL.EXE');
    end
    else
    %close all
invoke(Excel,'Quit');
delete(Excel);
    end
    
end
%% find drug for this session
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
    pName=fullfile(codeDir, 'getDrugsF.p');
    pNameFin=fullfile(drugDir,'getDrugsF.p');
    copyfile(pName,pNameFin)
    
    
fprintf('Thank you %s! Make sure you delete getDrugs.m and randomization.m scripts from the P drive \n and save them in your computer!\n',userS)
end 
clear
