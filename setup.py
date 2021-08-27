from cx_Freeze import *
import sys
includefiles=['icon.ico','attach.png','browse.png','clear.png','email.png','exit.png',
              'mic.png','music1.mp3','send.png','setting.png']
base=None
if sys.platform=="win32":
    base="Win32GUI"

shortcut_table=[
    ("DesktopShortcut",
     "DesktopFolder",
     "Email Sender",
     "TARGETDIR",
     "[TARGETDIR]\main.exe",
     None,
     None,
     None,
     None,
     None,
     None,
     "TARGETDIR",
     )
]
msi_data={"Shortcut":shortcut_table}

bdist_msi_options={'data':msi_data}
setup(
    version="0.0.1",
    description="Email Sender Application with Bulk Email Sender",
    author="uaahacker",
    name="Bulk Email Sender",
    options={'build_exe':{'include_files':includefiles},'bdist_msi':bdist_msi_options,},
    executables=[
        Executable(
            script="main.py",
            base=base,
            icon='icon.ico',
        )
    ]
)