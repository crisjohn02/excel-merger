# excel-merger
python excel merger
1.	Source file – Select the source file for merging. Example of this file is the demographics excel file. This file can be empty if you only want to extract data from the URL
2.	Excluded columns – Input the column names of the source file to be excluded in merging
3.	Base column – This column will be used to compare the target excel file. For example, you want to merge the demographics with ‘Respondent ID’ column matched with other excel column named ‘Sample ID’ – the program will align the rows based on the row value
4.	Target file – Refers to the target excel file. Example of this file is the Impress respondent level data.
5.	Based column – This column will be compared against the based column for the source file.
6.	Insert to index – Specify the column index where you want to put the appended data from the source file
7.	Extract data from URL – You can specify what data from URL query you want to add to your rows. Examples are brc, group, gender, etc. (data from fluent). URL: refers to the source column for your URL followed by the URL data index delimited by comma.
