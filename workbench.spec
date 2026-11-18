# -*- mode: python ; coding: utf-8 -*-
from pathlib import PurePosixPath

analysis = Analysis(
    ["main.py"],
    pathex=[SPECPATH],
    # Pillow supports Tk previews and pandas supports optional plotting. The
    # desktop app renders its own Qt widgets and processes images with Pillow.
    excludes=["tkinter", "_tkinter", "PIL.ImageTk", "PIL.AvifImagePlugin", "PIL._avif", "matplotlib", "scipy",
              "cv2", "PyQt5", "PySide6", "IPython", "pytest", "pandas.tests"],
)


def needed_qt_file(entry):
    """按 Qt 资源类型保留窗口所需 DLL、Windows 插件和中文翻译。"""
    # The only Qt rendering target is Windows. Image codecs for input files
    # come from Pillow. Keep Qt's Windows platform and native widget styles.
    parts = PurePosixPath(entry[0].replace("\\", "/")).parts
    if parts[:3] == ("PyQt6", "Qt6", "bin"):
        # Dependencies of discarded image plugins, plus the optional software
        # OpenGL backend. The application uses the QWidget raster painter.
        return parts[-1] not in {"Qt6Pdf.dll", "Qt6Svg.dll", "Qt6Network.dll", "opengl32sw.dll"}
    if parts[:3] == ("PyQt6", "Qt6", "plugins"):
        return parts[3] == "styles" or parts[3:] == ("platforms", "qwindows.dll")
    if parts[:3] == ("PyQt6", "Qt6", "translations"):
        return parts[-1] in {"qtbase_zh_CN.qm", "qtbase_zh_TW.qm"}
    return True


analysis.binaries = [entry for entry in analysis.binaries if needed_qt_file(entry)]
analysis.datas = [entry for entry in analysis.datas if needed_qt_file(entry)]
pyz = PYZ(analysis.pure)
exe = EXE(
    pyz, analysis.scripts, analysis.binaries, analysis.datas, [],
    name="归档台", console=False, upx=False, strip=False,
)
