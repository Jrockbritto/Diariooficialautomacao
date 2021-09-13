from cx_Freeze import setup, Executable

base = None


executables = [Executable("automation.py",
               base=base,
               icon="img/icon.ico",
               targetName="Automation"
               )]

packages = ['os','sys', 'logging']

options = {
    'build_exe': {    
        'packages':packages,
        'include_files': ['img/'],
        'includes': ["selenium", "webdriver_manager.chrome", "time", "xlsxwriter", "openpyxl", "datetime"],
        'build_exe': 'Automation',
        'excludes': ['tkinter', 'ctypes', 'html', 'pydoc_data', 'test', 'xmlrpc'] #otimizando script
    },
}

setup(
    name = "AutomaçãoDO",
    options = options,
    version = "1.0",
    description = 'Script que busca por documentos que possuam uma substring.',
    executables = executables
)
