from cx_Freeze import setup,Executable
import sys
import os

os.environ['TCL_LIBRARY'] = r'C:\\Users\\username\\AppData\\Local\\Programs\\Python\\Python37\\tcl\\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\\Users\\username\\AppData\\Local\\Programs\\Python\\Python37\\tcl\\tk8.6'

build_exe_options={"packages":['pandas','os','glob','numpy','openpyxl']} 

setup(
    name="ReportGenerator",
    version='0.1',
    description='Reporting',
    options={"build_exe":build_exe_options},
    executables=[Executable(script="reportgenerator.py", base = 'Console',icon="reporting.ico")])
