
REM ***************************************************************************
REM *** CreateReadUpdateDelete-FolderTextCsvExcelJson                       ***
REM ***************************************************************************

REM === �ߋ����O�̍폜 ===
DEL .\*.log

REM === �t�H���_�쐬 ===
CSCRIPT /E:jscript CrudFile.js CreateFolder .\tempFolder

REM === �e�L�X�g�t�@�C���쐬 ===
CSCRIPT /E:jscript CrudFile.js CreateText .\tempFolder\text1.txt

REM === CSV�t�@�C���쐬 ===
CSCRIPT /E:jscript CrudFile.js CreateCsv .\tempFolder\csv1.csv

REM === Excel�t�@�C���쐬 ===
CSCRIPT /E:jscript CrudFile.js CreateExcel .\tempFolder\excel1.xlsx

REM === Json�t�@�C���쐬 ===
CSCRIPT /E:jscript CrudFile.js CreateJson .\tempFolder\json1.json

REM === �e�L�X�g�t�@�C���Ǎ� ===
CSCRIPT /E:jscript CrudFile.js ReadText .\tempFolder\text1.txt

REM === CSV�t�@�C���Ǎ� ===
CSCRIPT /E:jscript CrudFile.js ReadCsv .\tempFolder\csv1.csv

REM === Excel�t�@�C���Ǎ� ===
CSCRIPT /E:jscript CrudFile.js ReadExcel .\tempFolder\excel1.xlsx

REM === Json�t�@�C���Ǎ� ===
CSCRIPT /E:jscript CrudFile.js ReadJson .\tempFolder\json1.json

REM === �e�L�X�g�t�@�C���X�V ===
CSCRIPT /E:jscript CrudFile.js UpdateText .\tempFolder\text1.txt

REM === CSV�t�@�C���X�V ===
CSCRIPT /E:jscript CrudFile.js UpdateCsv .\tempFolder\csv1.csv

REM === Excel�t�@�C���X�V ===
CSCRIPT /E:jscript CrudFile.js UpdateExcel .\tempFolder\excel1.xlsx

REM === �t�H���_�Ǎ� ===
CSCRIPT /E:jscript CrudFile.js ReadFolder .\tempFolder

REM === �t�H���_���� ===
CSCRIPT /E:jscript CrudFile.js CopyFolder .\tempFolder .\folder1

REM === �t�@�C���폜 ===
CSCRIPT /E:jscript CrudFile.js DeleteFile .\tempFolder\csv1.csv

REM === �t�H���_�폜 ===
CSCRIPT /E:jscript CrudFile.js DeleteFolder .\tempFolder

pause
