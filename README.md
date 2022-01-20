# VBA 100本knock

+ See https://excel-ubara.com/vba100/
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
wget https://github.com/knknkn1162/vba100_knock/releases/download/test/vba100_books.zip
unzip ./vba100_books.zip -d ./src
# install nkf and make in Windows. See in detail; https://github.com/knknkn1162/excel_vba_skeleton
Start-Process powershell -Verb runAs
choco source add -n kai2nenobu -s https://www.myget.org/F/kai2nenobu
choco install -y nkf make
```

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
