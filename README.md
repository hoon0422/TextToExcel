# Text_to_Excel

This program is a GUI program that converts text files containing test data of Acacia Communications Inc. to Excel
files. It is based on Python with version of more than 3.6 and uses PyQt5 for GUI.

Text files have a name with specified format for test data. Each data in each file will be combined with
a template Excel file, and will be saved as Excel file format. On combining, user can regard some values as 
missing value and erase them. 


## Installation & Execution
Download or clone this repository and run ```text_to_excel.bat``` file.

## Manual
1. Select a template file. The file must be an Excel file.
2. Select a sheet name for a template.
3. If you want to create a new Excel file which is the result of merging data in the text files,
write the range of data in each text file you want to merge in Excel range format.
4. Add text files for test data in ```Data List```
5. Write values regarded as missing values separated by space for each data.
6. Click button ```Set Save File Names```
7. Set names of newly created files which is the result of combining test data and a template.
If you wrote a range for merge, you can also insert the name of an Excel file created to merge data in
text files. The format of the name of created files is ```[serialnumber] [name user inserted] [date and time]```.
8. Click OK.
