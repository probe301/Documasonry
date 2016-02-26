
# import sys
import os
from cx_Freeze import setup, Executable



version = '2.0'
app_name = 'Documasonry'
icon = 'icon.ico'
description = "Documasonry - Multiple documents generating tool"
main_script = 'documasonry_gui.py'


include_files = ("build/msvcp100.dll",
                 "build/msvcr100.dll",
                 'documasonry_gui.ui',
                 'config.yaml',
                 # 'templates',
                 'icon.ico',
                 # 'doc',
                 )




build_path = os.getcwd() + '/build/{}_v{}/'.format(app_name, version)
build_exe_options = {
  "packages": ["os", ],
  "excludes": ["tkinter"],
  "includes": ["re", "atexit", ],
  "icon": icon,
  "build_exe": build_path,
  "include_files": include_files
}



setup(name=app_name + '.exe',
      version=version,
      description=description,
      options={"build_exe": build_exe_options},
      executables=[Executable(main_script, base="Win32GUI", icon=icon)])

os.remove(build_path + '/Qt5WebKit.dll')

pythoncom = "C:/Anaconda3/Lib/site-packages/pywin32_system32/pythoncom34.dll"
import shutil
shutil.copy(pythoncom, build_path + '/pythoncom34.dll')
# os.remove(build_path + '/icudt53.dll')  # 这个20m的文件不能删



