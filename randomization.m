function [randList]=randomization(N, Ndrugs)
% function that creates the randomization list for a placebo-controlled drug
% study with one drug and saves it in an excel file. Drug order is 
% counterbalanced across sessions and we can limit the number of maximum 
% consecutive repetitions of same drug allowed

%control for matlab repeatability
rng shuffle

% where the randomization list should be saved
outputFolder = '\\fileserver.dccn.nl\project\3024005.02\TestingDay\MedicationPreparation';
outputFolder='C:\Users\danpap\Documents\GitHub\RandList';

% name of output file
outFilename = fullfile(outputFolder,'randList.xlsx');

subNo = 1:N;
%maximum allowed consecutive repetitions
maxReps=4;
%loop until vector reaches repetition limit
numReps=5;

while any(numReps>maxReps)
    
index = randperm(N); 
%shuffle drugs with index
day1 = repmat(1:Ndrugs,1,N);day1=day1(index);
%count consecutive repetitions
numReps=diff([0 find(diff(day1)) numel(day1)]);

end

%reverse drug order, shuffled with same index
day2 = repmat(wrev(1:Ndrugs),1,N); day2=day2(index);
randList = [subNo' day1' day2'];

xlswrite(outFilename,randList)