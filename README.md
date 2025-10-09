## Invoice Manager

Programmes written to automate the process of signing invoice documents by quick generation of DOCX files based on a simple template.  
Written for use in the Polish language, it can be adjusted with just a few tweaks.

- Data Editor - manages data used for printing invoices
- Invoice Printer - used for printing selected data into a DOCX

Managed using [UV](https://docs.astral.sh/uv/)

Compiled using PyInstaller 6.8.0:
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