from cx_Freeze import setup, Executable
 
# Dependencies are automatically detected, but it might need
# fine tuning.
files = ['pdf.ico', 'config.ini']

build_options = { 'packages': [], 'excludes': [], 'include_files' : files}#'packages': [], 'excludes': [],
 
import sys
base = 'console' if sys.platform=='win32' else None
 
executables = [
    Executable('main.py', base=base, target_name = 'Stampinator', icon= "pdf.ico")
]
 
setup(name= 'Stampinator',
      version = '2',
      description = 'Eine Stempel App',
      options = {'build_exe': build_options},
      executables = executables)