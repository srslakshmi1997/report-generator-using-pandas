# report-generator-using-pandas
An application to generate a consolidated report from multiple excel files

### Overview :
This application pulls data from multiple excel files, consolidates all the data and prints the final report to another excel sheet. The location and name of the final excel sheet can be chosen by the end-user. This is a console-based application that can be run as windows executable.

### File Description :
**Files/Report - Shaun.xlsx :** A sample Excel file that serves as input for the code to generate a report. Multiple excel files of a different person can be placed along with this file to get a consolidated report.

**Files/Final_Report.xlsx :** This is the resultant Excel file that is generated on running reports on multiple files. This is the final output of the application. The consolidated report is written in this file. 

**Executable Setup.txt :** This file lists out the step to generate an executable application on this python program.

**XlsToCsv.vbs :** This vb script converts the data if any in .xlsx format to .csv format. This conversion enables the application to process the data effectively.

**download.log :** Unexpected error if any occurred, will be written in this file. This will help in the effective debugging process.

**reportgenerator.py :** The python file contains the logic to call the vb script to convert the files, gather data from all the files under the 'Files' folder, generate a report and write it back to the target excel file.

**reporting.ico :** A sample icon file for the executable.

**setup.py :** This python file will build and generate an executable file. Even a  windows installer can be created which is distributable.

