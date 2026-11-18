# -*- mode: python ; coding: utf-8 -*-
# The pinned customtkinter and _tkinter hooks collect fonts, themes and Tcl/Tk.
analysis = Analysis(
    ["main_customtkinter.py"],
    pathex=[SPECPATH],
    datas=[("assets/workbench-header.png", "assets")],
    excludes=["batch_ui", "rule_forms", "PyQt6", "PyQt5", "PySide6", "PySide2",
              "matplotlib", "scipy", "cv2", "IPython", "pytest", "pandas.tests",
              "PIL.AvifImagePlugin", "PIL._avif"],
)
pyz = PYZ(analysis.pure)
exe = EXE(
    pyz, analysis.scripts, analysis.binaries, analysis.datas, [],
    name="归档台-CustomTkinter", console=False, upx=False, strip=False,
)
