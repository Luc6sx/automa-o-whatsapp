import sys
from cx_Freeze import setup, Executable


#Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {
    "packages": ["os"],
    "includes": ["PySimpleGUI", "openpyxl","webbrowser", "pyautogui"]  # Dependências que seu bot pode usar
}


#GUI applications require a different base on Windows (the default is for
#a console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="bot whatsapp",
    version="0.1",
    description="manda menssagem automaquita para pessoas de contato no exell",
    options={"build_exe": build_exe_options},
    executables=[Executable("bot.py", base=base)]
)