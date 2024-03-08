## Tax Printer (and Data Editor)

Programmes written to automate the process of signing documents by qucik generation of docx files based on a simple template
Written for use in Polish language, can be adjusted with just a few tweaks

Compiled using pyinstaller:
```
pyinstaller {file} ^
  --clean ^
  --onefile ^
  --noconsole ^
  --noconfirm ^
  --distpath .\\ ^
  --workpath .\\{dump_folder} ^
  --specpath .\\{dump_folder} ^
  --hidden-import babel.numbers ^
  --icon ..\\{icon}
```

Made with:
- python
- TKinter
