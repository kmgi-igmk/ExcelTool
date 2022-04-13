# About ExcelTool

## Requirements

- Python 3.10.4

## How you develop

**1. make sure to install python(3.10.4 or higher)**
**2. create venv at working directory if necessary**
i.e. `python -m venv pyexe`
apply `--clear` option if venv should be initialised.
**3. activate it**
`pyexe\Scripts\activate.bat`
**4. install dependencies**
install modules by pip on your own
`pip install pandas openpyxl pyinstaller`
or
install modules by requirements.txt
`python -m pip install -r requirements.txt`
※you can create requirements.txt by below command
`python -m pip freeze > requirements.txt`

## How you build

**1. make sure that \*.spec files exist**
`OneDir.spec`, `OneFile.spec`, `Scatter.spec`
**2. issue build command**
select one of these .spec file.
`pyinstaller {}.spec`
※remove temporary files and dir such as *build/* when `--clean` option is applied.
or
you can build it on your own.
`pyinstaller src\App.py --noconsole --icon=src\resources\icon.ico [ --onedir | --onefile | <Nothing> ] --name=ExcelTool --exclude sqlite3 --add-date=src\resources\icon.ico;.\resources\`
if build is successfully completed, ExcelTool.spec will create. Then you should add below code to that spec file in order to remove **'MKL'** module, which size is larger than others, so that product should be lighter and faster.

```spec
・・・・・・
a = Analysis(['src\\App.py'],
・・・・・・
pyz = PYZ(a.pure,・・・

# ~~~~~ from here to... ~~~~~~~
Key = ['mkl']
def remove_from_list(input, keys):
    outlist = []
    for item in input:
        name, _, _ = item
        flag = 0
        for key_word in keys:
            if name.find(key_word) > -1:
                flag = 1
        if flag != 1:
            outlist.append(item)
    return outlist
a.binaries = remove_from_list(a.binaries, Key)
# ~~~~~ here ~~~~~~~~

exe = EXE(pyz,・・・
・・・・・・・
```

once updated .spec, issue below
`pyinstaller {filename}.spec --clean`
**3. give it to someone, maybe zip**

## Tips

### icon for .exe does not appear on explorer

close all explorer and then try `windows key` + `r` and type below.
`ie4uinit.exe -show`
