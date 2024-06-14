
REM ***************************************************************************
REM *** CreateReadUpdateDelete-FolderTextCsvExcelJson                       ***
REM ***************************************************************************

REM === 過去ログの削除 ===
DEL .\*.log

REM === フォルダ作成 ===
CSCRIPT /E:jscript CrudFile.js CreateFolder .\tempFolder

REM === テキストファイル作成 ===
CSCRIPT /E:jscript CrudFile.js CreateText .\tempFolder\text1.txt

REM === CSVファイル作成 ===
CSCRIPT /E:jscript CrudFile.js CreateCsv .\tempFolder\csv1.csv

REM === Excelファイル作成 ===
CSCRIPT /E:jscript CrudFile.js CreateExcel .\tempFolder\excel1.xlsx

REM === Jsonファイル作成 ===
CSCRIPT /E:jscript CrudFile.js CreateJson .\tempFolder\json1.json

REM === テキストファイル読込 ===
CSCRIPT /E:jscript CrudFile.js ReadText .\tempFolder\text1.txt

REM === CSVファイル読込 ===
CSCRIPT /E:jscript CrudFile.js ReadCsv .\tempFolder\csv1.csv

REM === Excelファイル読込 ===
CSCRIPT /E:jscript CrudFile.js ReadExcel .\tempFolder\excel1.xlsx

REM === Jsonファイル読込 ===
CSCRIPT /E:jscript CrudFile.js ReadJson .\tempFolder\json1.json

REM === テキストファイル更新 ===
CSCRIPT /E:jscript CrudFile.js UpdateText .\tempFolder\text1.txt

REM === CSVファイル更新 ===
CSCRIPT /E:jscript CrudFile.js UpdateCsv .\tempFolder\csv1.csv

REM === Excelファイル更新 ===
CSCRIPT /E:jscript CrudFile.js UpdateExcel .\tempFolder\excel1.xlsx

REM === フォルダ読込 ===
CSCRIPT /E:jscript CrudFile.js ReadFolder .\tempFolder

REM === フォルダ複写 ===
CSCRIPT /E:jscript CrudFile.js CopyFolder .\tempFolder .\folder1

REM === ファイル削除 ===
CSCRIPT /E:jscript CrudFile.js DeleteFile .\tempFolder\csv1.csv

REM === フォルダ削除 ===
CSCRIPT /E:jscript CrudFile.js DeleteFolder .\tempFolder

pause
