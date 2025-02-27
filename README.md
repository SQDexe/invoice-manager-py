## Tax Printer (and Data Editor)

Programmes written to automate the process of signing documents by qucik generation of docx files based on a simple template. Written for use in Polish language, can be adjusted with just a few tweaks

Compiled using PyInstaller:
```
pyinstaller {file} ^
  --name "{name}" ^
  --icon ..\\{icon} ^
  --optimize 2 ^
  --clean ^
  --onefile ^
  --noconsole ^
  --noconfirm ^
  --distpath .\\ ^
  --workpath .\\{dump_folder} ^
  --specpath .\\{dump_folder} ^
  --hidden-import babel.numbers
```

Most prominent dependencies:
- Python 3.12.4 ( TKinter 8.6.13, PyInstaller 6.8.0 )
