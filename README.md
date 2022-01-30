# VBA 100本knock

+ See https://excel-ubara.com/vba100/
    + skip [ex068](https://excel-ubara.com/vba100/VBA100_068.html) because this problem is about form-control
+ PowerShell version: [knknkn1162/vba100_knock_ps](https://github.com/knknkn1162/vba100_knock_ps)

## How to import/run macros

Use [knknkn1162/excel_vba_skeleton](https://github.com/knknkn1162/excel_vba_skeleton) tools:

1: git clone [knknkn1162/excel_vba_skeleton](https://github.com/knknkn1162/excel_vba_skeleton#macos)

2: Download xlsm books (Each book doesn't contiain macro)

3: Git clone this repo

4-1: \[Windows\] Install `nkf` and `make`. See in detail; https://github.com/knknkn1162/excel_vba_skeleton#prerequisites

4-2: \[Mac OS\] See the link; https://github.com/knknkn1162/excel_vba_skeleton#macos

```sh
git clone --recursive https://github.com/knknkn1162/excel_vba_skeleton ./proj
cd proj
# get this repo
git clone https://github.com/knknkn1162/vba100_knock ./src
# get xlsm books
wget https://github.com/knknkn1162/vba100_knock/releases/download/books/books.zip
unzip ./vba100_books.zip -d ./src
# install nkf and make in Windows. See in detail; https://github.com/knknkn1162/excel_vba_skeleton
Start-Process powershell -Verb runAs
choco source add -n kai2nenobu -s https://www.myget.org/F/kai2nenobu
choco install -y nkf make
```

4-3: load references:

+ ex071: load "Microsoft PowerPoint xx.x ObjectLibrary"
    + excel > Alt+F11(Open VBE) > tool > References > Check "Microsoft PowerPoint xx.x Object Library"
+ ex079: load "Microsoft Word xx.x ObjectLibrary"
    + excel > Alt+F11(Open VBE) > tool > References > Check "Microsoft Word xx.x Object Library"
+ ex097, ex098: load "Microsoft ActiveX Data Objects 6.1 Library"
    + excel > Alt+F11(Open VBE) > tool > References > Check "Microsoft ActiveX Data Objects 6.1 Library"
+ ex100: use
    + download [Serenium Basic](https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0)
    + download [Chrome Driver](https://sites.google.com/chromium.org/driver/)
    + download [Microsoft .NET Framework 3.5](https://www.microsoft.com/ja-jp/download/details.aspx?id=25150)
    + For more information, see the [link(japanese)](https://excel-ubara.com/excelvba4/EXCEL_VBA_401.html)

5: Type make commands below;

You can import/run macro by bash or powershell:

```sh
# import all macros into books
make import-all # or `make import XLSM=ex001`
# launch excel application and run macro automatically.
## choose xlsm book as you like.
make run XLSM=ex001
```

## directories

+ omit ex003~ex100 for simplicity

```
proj
├── books
│   ├── ex001.xlsm
│   ├── ex002.xlsm
├── scripts
├── src(this repo)
│   ├── ex001/Module1.bas
│   ├── ex002/Module1.bas
├── templates
└── vbac
```
