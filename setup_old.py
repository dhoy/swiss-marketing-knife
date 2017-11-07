from cx_Freeze import setup, Executable

exe=Executable(
     script='ftp_exact_target.py',
     base="Win32Gui",
     icon="image/swiss-marketing-knife.ico",
     targetName="Swiss Marketing Knife.exe"
     )
includefiles=['image/','download/', 'upload/', 'resource_rc.py', 'C:\Python34\Lib\site-packages\PyQt5\libEGL.dll']
includes=['decimal', 'atexit']
excludes=['Tkinter']
packages=['ui']
setup(
     version = "0.1",
     description = "Tool to help James match email address by DMCID, and now much, much more!",
     author = "David Hoy",
     name = "Swiss Marketing Knife",
     options = {'build_exe': {'excludes':excludes,'packages':packages,'include_files':includefiles,'includes':includes}},
     executables = [exe]
     )