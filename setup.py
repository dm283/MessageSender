#  Для формирование exe-файлов:  python setup.py build

from cx_Freeze import setup, Executable

includefiles = ['config.ini', 'rec-k.txt', 'README.md']
includes = [] 
excludes = []

base = "Win32GUI"

setup(
 name=['mSender'],
 version = '1.0',
 description = 'mSender',
 options = {'build_exe':   {'excludes':excludes,'include_files':includefiles, 'includes':includes}},
 executables = [
    Executable('mSender.py', base=base), 
    Executable('mSenderConsole.py'),  #  base=base не указывается, при запуске не появляется окно cmd
    Executable('mSenderAdministration.py'), 
    Executable('mSenderCreateMsg.py', base=base)]
)
