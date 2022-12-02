from cx_Freeze import setup, Executable

setup(name='Nutricash',
      version='2.0',
      description='Nutricash',
      options={'build_exe': {'packages': ['numpy','pandas','selenium','BeautifulSoup','xlrd','dateutil']}},
      executables=[Executable(script='Main.py', base="Win32GUI")]
)