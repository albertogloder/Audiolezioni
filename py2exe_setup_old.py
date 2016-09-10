import os
import os.path
from distutils.core import setup
import py2exe
import sys
sys.argv.append('py2exe')

options = dict(
    bundle_files=1,
    optimize=2,
    compressed=True,
    excludes=['_ssl', 'pyreadline', 'difflib', 'doctest', 'locale', 'optparse', 'pickle', 'email', 'calendar'],
    dll_excludes=['msvcr71.dll'],
)

setup(
    version='1.0',
    description="Questo programma stampa la lista delle lezioni da tenere d'occhio.",
    author='Alberto Gloder',

    options={'py2exe': options},
    windows=[{'script': "futuristic_audio_lessons.py", 'icon_resources': [(1, 'icon.ico')]}],
    data_files=["audio_lezioni_semestre.xls", "phantomjs.exe"],
    zipfile=None,
)