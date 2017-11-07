from cx_Freeze import setup, Executable

exe=Executable(
     script='ftp_exact_target.py',
     base="Win32Gui",
     icon="image/swiss-marketing-knife.ico",
     targetName="Swiss Marketing Tool.exe",
     #shortcutName="Remove From Mailing List",
     #shortcutDir="ProgramMenuFolder",
     )

shortcut_table = [
    ("DesktopShortcut",        # Shortcut
     "DesktopFolder",          # Directory_
     "Swiss Marketing Tool",           # Name
     "TARGETDIR",              # Component_
     "[TARGETDIR]Swiss Marketing Tool.exe",# Target
     None,                     # Arguments
     None,                     # Description
     None,                     # Hotkey
     None,                     # Icon
     None,                     # IconIndex
     None,                     # ShowCmd
     'TARGETDIR'               # WkDir
     ),
    ("ProgramShortcut",        # Shortcut
     "StartMenuFolder",          # Directory_
     "Swiss Marketing Tool",           # Name
     "TARGETDIR",              # Component_
     "[TARGETDIR]Swiss Marketing Tool.exe",# Target
     None,                     # Arguments
     None,                     # Description
     None,                     # Hotkey
     None,                     # Icon
     None,                     # IconIndex
     None,                     # ShowCmd
     'TARGETDIR'               # WkDir
     )
    ]

msi_data = {"Shortcut": shortcut_table}

build_exe_options = {'packages':['ui'], 'include_files':['image','download/', 'upload/', 'resource_rc.py', 'C:\Python34\Lib\site-packages\PyQt5\libEGL.dll'], 
                     'includes':['decimal', 'atexit'], 'excludes':['Tkinter'], 'include_msvcr': True}

bdist_msi_options = {
    #GUID -- use generator online for now.
    'upgrade_code': '{390b43e0-5504-4e5e-afb2-f8ed9ad7cf54}',
    'add_to_path': False,
    'data': msi_data,
    #'initial_target_dir': r'[Program Files] \%s\%s' % ('test', 'test'),
    }
setup(
     version = "2.14",
     description = "Tool to help James match email address by DMCID, and now much, much more!",
     author = "David Hoy",
     name = "Swiss Marketing Tool",
     options = {'build_exe': build_exe_options, 'bdist_msi': bdist_msi_options},
     executables = [exe]
     )