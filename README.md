# VBA 100本knock

+ See https://excel-ubara.com/vba100/

## How to import/run macros

1. git clone tool; https://github.com/knknkn1162/excel_vba_skeleton

2. get xlsm books
3. git clone this repo

```sh
git clone --recursive https://github.com/knknkn1162/excel_vba_skeleton ./proj
cd proj
# get this repo
git clone https://github.com/knknkn1162/vba100_knock ./vba100
make import-vba100
# get xlsm books
wget https://github.com/knknkn1162/vba100_knock/releases/download/test/vba100_xlsms.zip
unzip ./vba100_xlsm.zip
```

You can run macro by bash or powershell:

```sh
make run XLSM=./vba100/xlsms/ex001.xlsm
```
