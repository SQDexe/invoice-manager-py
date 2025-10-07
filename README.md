## Invoice Manager

Programmes written to automate the process of signing invoice documents by qucik generation of DOCX files based on a simple template.
Written for use in Polish language, can be adjusted with just a few tweaks.

- Data Editor - manages data used for prinitng invoices
- Invoice Printer - used for prinitng selected data into a DOCX

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
