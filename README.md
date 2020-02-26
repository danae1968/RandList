# RandList
Instructions for external researcher that generates randomization list for pharmacological study.
Purpose: 
A. Generate Randomization List for pharmacological study
B. Generate Script that provides drug for specific session/subject combination and documents every time the script was ran. To ensure double-blindness in a study that medication needs to be prepared for every session.


# Step 1: 
Go to the randomization.m script in P:\3024005.02\Taskcode\RandList and **add your selected password** in the indicated section (type password here).
Run the randomization.m script: randomization(N, Ndrugs). N is the number of subjects, Ndrugs the number of drugs. 
In our case you should run: **randomization(30,2)**. The output is "randList.xlsx" file that should be an N by ndrugs+1 matrix (30x3). The first column should be the subject number and the other two the sessions.
Check that it is the case. 

# Step 2
Go to the script getDrugs.m. **Copy paste the password from the previous script** in the section %% ADD EXCEL PASSWORD HERE and **run the script**. 

# Step 3
Remove the randomization.m, the getDrugs.m script and the getDrugsF.m scripts from the P drive and store them in a safe directory of your computer. Very important that they are not lost. 

# Step 4
Run getDrugsF.p to see if it works!
