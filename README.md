# RandList
Instructions for external researcher that generates randomization list for pharmacological study.
Purpose: 
A. Generate Randomization List for pharmacological study
B. Generate Script that provides drug for specific session/subject combination and documents every time the script was ran. To ensure double-blindness in a study that medication needs to be prepared for every session.


# Step 1: 
Go to the randomization.m script and **modify the outputFolder** variable with the directory you want to save the list in. Run the randomization.m script: randomization(N, Ndrugs). N is the number of subjects, Ndrugs the number of drugs. 
In our case you should run: **randomization(30,2)**. The output is "randList.xlsx" file that should be an N by ndrugs+1 matrix (30x3). The first column should be the subject number and the other two the sessions.

# Step 2
Pick a password and **password protect the excel file: randList.xlsx**. Then go to the script getDrugs.m. **Adapt the directories** of this section:%% fill in here relevant pathways to match the directory in which you saved the file and **run the script**. 

# Step 3
Store getDrugsF.m script that contains the password for the randomization list in a save directory of your computer. Very important that it is not lost. 

# Step 4
Provide the researchers with getDrugs.p, randList.xlsx and runLog.xlsx files.
