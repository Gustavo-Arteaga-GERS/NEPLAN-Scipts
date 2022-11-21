# NEPLAN-Scipts
The scripting functionality of Neplan offers an interface to the software via programming

## Get Started 

To execute the contingency analysis the following sequence must be followed: 

1) Set the local Path of your PC where the project is cloned (see line 25 of the file named scriptSaveDataCSV.csx).

2) in the file named "elementsContingency.csv" list the elements that you want to open/close in each study.

3) in the file named "elementsForAnalysis.csv" list the name of the elements for which you want to obtain the results of each contingency (you must specify the name followed by a comma and the type of element, e.g. Armenia - Hermosa 1 230,line). 

4) Execute the analysis from the Run script option of the Special Analysis tab in the NEPLAN application. 

5) Once the analysis is finished, the results obtained will be automatically saved in .csv files, the file name is the date followed by the exact sheet of each analysis. 
