from cx_Freeze import setup, Executable
import os
import sys

base = None


executables = [Executable("automation.py",
               base=base,
               icon="img/icon.ico",
               targetName="Automation"
               )]

packages = ["os",'sys']

options = {
    'build_exe': {    
        'packages':packages,
        'include_files': ['img/'],
        'includes': ["selenium", "webdriver_manager.chrome", "time", "xlsxwriter", "openpyxl", "datetime"],
        'build_exe': 'Automation',
        'excludes': ['tkinter']
    },
}

setup(
    name = "AutomaçãoDO",
    options = options,
    version = "0.5",
    description = 'Script que busca por documentos que possuam uma substring.',
    executables = executables
)
