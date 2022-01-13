# VBA 100本knock

+ See https://excel-ubara.com/vba100/

## How to import/run macros

Use [knknkn1162/excel_vba_skeleton](https://github.com/knknkn1162/excel_vba_skeleton) tools:

1: git clone [knknkn1162/excel_vba_skeleton](https://github.com/knknkn1162/excel_vba_skeleton#macos)

2: Get xlsm books

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
choco source add -n kai2nenobu -s https://www.myget.org/F/kai2nenobu
choco install -y nkf make
```

5: Type make commands below;

You can import/run macro by bash or powershell:

```sh
# import all macros into books
make import-vba100 # or `make run XLSM=./vba100/xlsms/ex001.xlsm`
# launch excel application and run macro automatically.
## choose xlsm book as you like.
make run XLSM=./vba100/xlsms/ex001.xlsm
```
