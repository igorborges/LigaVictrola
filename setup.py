from distutils.core import setup
import py2exe

setup(
    console = [
        {
            "script": "Menu V.I.C.artola.py",
            "icon_resources": [(1, "icon.ico")]
        }
    ],
)