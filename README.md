# pythonExcel
An example on how to use the python library openpyxl to manipulate excel files.
The script shows how to manipulate excel files with the library openpyxl v. 2.4.8.

It looks for a excel file named ExcelFile. In this files there are differen sheets, each showing the correspondance of elements belonging to different classes:
- correspondance between Level 1 elements and Level 2 elements
- correspondance between Level 2 elements and Level 3 elements
- correspondance between Level 3 elements and Level 4, 5, 6 elements
- correspondance between Level 2 elements and Level 6 elements
- correspondance between Level 3 elements and Level 6 elements.

The script takes this file and create a new file with the dependency in one single sheet. The output file is named outputExcelFile.

To run the script execute the command 'python excelPython.py'

Alternatively you can launch the script gui.py with the command 'python gui.py' and press the button 'Create New Match. The gui.py is a simple gui to search and inspect the elements.
