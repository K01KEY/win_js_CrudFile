/**
 * [Create],[Read],[Update],[Delete]-[Folder],[text],[csv],[excel]-[json]
 */
// このファイルはSJISで記述されています。

var fso = new ActiveXObject("Scripting.FileSystemObject");
var wsh = new ActiveXObject("WScript.Shell");

//-----------------------------------------------------------------------------
var global = Function('return this')();
if (!global.JSON) {
  global.JSON = {
    parse: function(sJSON) { return eval('(' + sJSON + ')'); },
    stringify: (function () {
      var toString = Object.prototype.toString;
      var hasOwnProperty = Object.prototype.hasOwnProperty;
      var isArray = Array.isArray || function (a) { return toString.call(a) === '[object Array]'; };
      var escMap = {'"': '\\"', '\\': '\\\\', '\b': '\\b', '\f': '\\f', '\n': '\\n', '\r': '\\r', '\t': '\\t'};
      var escFunc = function (m) { return escMap[m] || '\\u' + (m.charCodeAt(0) + 0x10000).toString(16).substr(1); };
      var escRE = /[\\"\u0000-\u001F\u2028\u2029]/g;
      return function stringify(value) {
        if (value == null) {
          return 'null';
        } else if (typeof value === 'number') {
          return isFinite(value) ? value.toString() : 'null';
        } else if (typeof value === 'boolean') {
          return value.toString();
        } else if (typeof value === 'object') {
          if (typeof value.toJSON === 'function') {
            return stringify(value.toJSON());
          } else if (isArray(value)) {
            var res = '[';
            for (var i = 0; i < value.length; i++)
              res += (i ? ', ' : '') + stringify(value[i]);
            return res + ']';
          } else if (toString.call(value) === '[object Object]') {
            var tmp = [];
            for (var k in value) {
              // in case "hasOwnProperty" has been shadowed
              if (hasOwnProperty.call(value, k))
                tmp.push(stringify(k) + ': ' + stringify(value[k]));
            }
            return '{' + tmp.join(', ') + '}';
          }
        }
        return '"' + value.toString().replace(escRE, escFunc) + '"';
      };
    })()
  };
}

//-----------------------------------------------------------------------------
function formatDate0 (date, format) {
  format = format.replace(/yyyy/g, date.getFullYear());
  format = format.replace(/MM/g, ('0' + (date.getMonth() + 1)).slice(-2));
  format = format.replace(/dd/g, ('0' + date.getDate()).slice(-2));
  format = format.replace(/HH/g, ('0' + date.getHours()).slice(-2));
  format = format.replace(/mm/g, ('0' + date.getMinutes()).slice(-2));
  format = format.replace(/ss/g, ('0' + date.getSeconds()).slice(-2));
  format = format.replace(/SSS/g, ('00' + date.getMilliseconds()).slice(-3));
  return format;
};
function formatDate (date, format) {
  format = format.replace(/yyyy/g, date.getFullYear());
  format = format.replace(/M/g, (date.getMonth() + 1));
  format = format.replace(/d/g, (date.getDate()));
  format = format.replace(/H/g, (date.getHours()));
  format = format.replace(/m/g, (date.getMinutes()));
  format = format.replace(/s/g, (date.getSeconds()));
  format = format.replace(/S/g, (date.getMilliseconds()));
  return format;
};

//-----------------------------------------------------------------------------
function withSubFolderNames(path, func) {
  var folder = fso.getFolder(path);
  var em = new Enumerator(folder.SubFolders);
  for(em.moveFirst();!em.atEnd();em.moveNext()) {
    func(em.item().Name);
  }
}

//-----------------------------------------------------------------------------
function withFileNames(path, func) {
  var folder = fso.getFolder(path);
  var em = new Enumerator(folder.Files);
  for(em.moveFirst();!em.atEnd();em.moveNext()) {
    func(em.item().Name);
  }
}

//-----------------------------------------------------------------------------
function withReadFile(filename, func) {
  var file = fso.OpenTextFile(filename, 1, false);
  try {
    func(file);
  } finally {
    file.Close();
  }
}

//-----------------------------------------------------------------------------
function withEachLine(filename, func) {
  withReadFile(filename, function (file) {
    while (!file.AtEndOfStream) {
      func(file.ReadLine());
    }
  });
}

//-----------------------------------------------------------------------------
function withAllLine(filename, func) {
  withReadFile(filename, function (file) {
    func(file.ReadAll());
  });
}

//-----------------------------------------------------------------------------
function withWriteFile(filename, func) {
  var file = fso.OpenTextFile(filename, 2, true);
  try {
    func(file);
  } finally {
    file.Close();
  }
}

//-----------------------------------------------------------------------------
// EXCEL操作用関数
function withExcel(visible, func) {
  var excel = new ActiveXObject("Excel.Application");
  excel.Visible = visible;
  excel.DisplayAlerts = false;
  try {
    func(excel);
  } finally {
    excel.Quit();
  }
}

//-----------------------------------------------------------------------------
function withWorkbook(filename, visible, readonly, func) {
  withExcel(visible, function (excel) {
    var workbook = excel.Workbooks.Open(filename, 0, readonly);
    try {
      func(workbook);
    } finally {
      workbook.Close();
    }
  });
}

//-----------------------------------------------------------------------------
function newWorkbook(filename, visible, func) {
  withExcel(visible, function (excel) {
    var workbook = excel.Workbooks.Add();
    try {
      func(workbook);
    } finally {
      workbook.SaveAs(fso.getAbsolutePathName(filename));
      workbook.Close();
    }
  });
}

//=============================================================================
// 主処理
//=============================================================================
/**
 * (1)初期処理
 */
//(1)-1:エラーコード初期化
var errCode = 0;

//(1)-2:現在日時を取得する
var wrkTimeStamp = getTimeStamp();

//(1)-3:動作環境のパスを取得する
var curPath = wsh.CurrentDirectory;
var appPath = fso.getParentFolderName(WScript.ScriptFullName); //このスクリプトのパス
var appName = fso.GetBaseName(WScript.ScriptFullName);         //このスクリプトの名前(拡張子なし)

//(1)-4:アプリケーションログのオープン
var appLogName = fso.BuildPath(appPath,appName + "_" + wrkTimeStamp + ".log");
var appLogFile = fso.OpenTextFile(appLogName,8,true,-1); //ファイル名,追加,新規作成,文字コード(UNI)
appLogFile.WriteLine("--- APP START ---");

//(1)-5:引数を取得する
var objArgs = WScript.Arguments;

//(1)-6:引数の妥当性をチェックする(ファイルレベル)
var argCmd,argFile1,argFile2;

switch(objArgs.length){
  case 0,1:
    //"引数が足りません。"
    WScript.Quit(9);
    break;
  case 2:
    argCmd   = objArgs(0);
    argFile1 = objArgs(1);
    break;
  case 3:
    argCmd   = objArgs(0);
    argFile1 = objArgs(1);
    argFile2 = objArgs(2);
    break;
  default:
    //"引数が多すぎます。"
    WScript.Quit(9);
}

/**
 * (2)メイン処理
 */
switch(argCmd.toLowerCase()){  //文字列を小文字に合わせて判定
  case "CreateFolder".toLowerCase():
    rtnCode = CreateFolder(argFile1);
    break;
  case "ReadFolder".toLowerCase():
    rtnCode = ReadFolder(argFile1);
    break;
  case "CopyFolder".toLowerCase():
    rtnCode = CopyFolder(argFile1, argFile2);
    break;
  case "DeleteFile".toLowerCase():
    rtnCode = DeleteFile(argFile1);
    break;
  case "DeleteFolder".toLowerCase():
    rtnCode = DeleteFolder(argFile1);
    break;
  case "CreateText".toLowerCase():
    rtnCode = CreateText(argFile1);
    break;
  case ("ReadText").toLowerCase():
    rtnCode = ReadText(argFile1);
    break;
  case ("ReadTextAll").toLowerCase():
    rtnCode = ReadTextAll(argFile1);
    break;
  case ("UpdateText").toLowerCase():
    rtnCode = UpdateText(argFile1);
    break;
  case ("CreateCsv").toLowerCase():
    rtnCode = CreateCsv(argFile1);
    break;
  case ("ReadCsv").toLowerCase():
    rtnCode = ReadCsv(argFile1);
    break;
  case ("UpdateCsv").toLowerCase():
    rtnCode = UpdateCsv(argFile1);
    break;
  case ("CreateExcel").toLowerCase():
    rtnCode = CreateExcel(argFile1);
    break;
  case ("ReadExcel").toLowerCase():
    rtnCode = ReadExcel(argFile1);
    break;
  case ("ReadExcelSheets").toLowerCase():
    rtnCode = ReadExcelSheets(argFile1);
    break;
  case ("UpdateExcel").toLowerCase():
    rtnCode = UpdateExcel(argFile1);
    break;
  case ("CreateJson").toLowerCase():
    rtnCode = CreateJson(argFile1);
    break;
  case ("ReadJson").toLowerCase():
    rtnCode = ReadJson(argFile1);
    break;
  case ("ALL").toLowerCase():
    rtnCode = CreateText(argFile1);
    break;
  default:
    WScript.Echo("CMD notfound!");
}

/**
 * (9)終了処理
 */
//(9)-1:エラー発生の判定

//(1)-2:アプリケーションログのクローズ
appLogFile.WriteLine("--- APP  END  ---");
appLogFile.Close();

//-----------------------------------------------------------------------------
function getTimeStamp() {
  var wrkNow = new Date();
  var yyyymmddhhmmss = wrkNow.getFullYear()+
	  ( "0"+( wrkNow.getMonth()+1 ) ).slice(-2)+
	  ( "0"+wrkNow.getDate() ).slice(-2)+
    ( "0"+wrkNow.getHours() ).slice(-2)+
    ( "0"+wrkNow.getMinutes() ).slice(-2)+
    ( "0"+wrkNow.getSeconds() ).slice(-2);
  return yyyymmddhhmmss;
}

//-----------------------------------------------------------------------------
function CreateFolder(iPath) {
  appLogFile.WriteLine("--- CreateFolder START ---");
  if (fso.FileExists(iPath)) {
    appLogFile.WriteLine("既にフォルダが存在します=" + iPath);
    return 9;
  }
  fso.CreateFolder(iPath);
  appLogFile.WriteLine("--- CreateFolder  END  ---");
}

//-----------------------------------------------------------------------------
function ReadFolder(iPath) {
  appLogFile.WriteLine("--- ReadFolder START ---");
  if (!fso.FolderExists(iPath)) {
    appLogFile.WriteLine("フォルダーが存在しません=" + iPath);
    return 9;
  }
  withSubFolderNames(iPath,function (folderName) {
    appLogFile.WriteLine("folderName=" + folderName);
    WScript.Echo("folderName=" + folderName);
  });
  withFileNames(iPath,function (fileName) {
    appLogFile.WriteLine("FileName=" + fileName);
    WScript.Echo("FileName=" + fileName);
  });
  appLogFile.WriteLine("--- ReadFolder  END  ---");
}

//-----------------------------------------------------------------------------
function CopyFolder(iPath1, iPath2) {
  appLogFile.WriteLine("--- CopyFolder START ---");
  if (!fso.FolderExists(iPath1)) {
    appLogFile.WriteLine("フォルダーが存在しません=" + iPath1);
    return 9;
  }
  try {
    fso.CopyFolder(iPath1, iPath2, true)  // 上書きコピーする
  } catch(e) {
    WScript.Echo("フォルダをコピーできませんでした");
    return 9;
  }
  appLogFile.WriteLine("--- CopyFolder  END  ---");
}

//-----------------------------------------------------------------------------
function DeleteFile(iFileName) {
  appLogFile.WriteLine("--- DeleteFile START ---");
  if (!fso.FileExists(iFileName)) {
    appLogFile.WriteLine("ファイルが存在しません=" + iFileName);
    return 9;
  }
  try {
    fso.DeleteFile(iFileName, false)  // 読み取り専用も削除
  } catch(e) {
    WScript.Echo("ファイルを削除できませんでした");
    return 9;
  }
  appLogFile.WriteLine("--- DeleteFile  END  ---");
}

//-----------------------------------------------------------------------------
function DeleteFolder(iPath) {
  appLogFile.WriteLine("--- DeleteFolder START ---");
  if (!fso.FolderExists(iPath)) {
    appLogFile.WriteLine("フォルダーが存在しません=" + iPath);
    return 9;
  }
  try {
    fso.DeleteFolder(iPath, false)  // 読み取り専用も削除
  } catch(e) {
    WScript.Echo("フォルダを削除できませんでした");
    return 9;
  }
  appLogFile.WriteLine("--- DeleteFolder  END  ---");
}

//-----------------------------------------------------------------------------
function CreateText(oFileName) {
  appLogFile.WriteLine("--- CreateText START ---");
  withWriteFile(oFileName,function (file) {
    for(var i=0;i<100;i++) {
      file.WriteLine("DATA" + i);
    }
  });
  appLogFile.WriteLine("--- CreateText  END  ---");
}

//-----------------------------------------------------------------------------
function ReadText(iFileName) {
  appLogFile.WriteLine("--- ReadText START ---");
  if (!fso.FileExists(iFileName)) {
    appLogFile.WriteLine("ファイルが存在しません=" + iFileName);
    return 9;
  }
  withEachLine(iFileName,function (line) {
    WScript.Echo(line);
  });
  appLogFile.WriteLine("--- ReadText  END  ---");
}

//-----------------------------------------------------------------------------
function ReadTextAll(iFileName) {
  appLogFile.WriteLine("--- ReadTextAll START ---");
  if (!fso.FileExists(iFileName)) {
    appLogFile.WriteLine("ファイルが存在しません=" + iFileName);
    return 9;
  }
  withAllLine(iFileName,function (buf) {
    WScript.Echo(buf);
  });
  appLogFile.WriteLine("--- ReadTextAll  END  ---");
}

//-----------------------------------------------------------------------------
function UpdateText(iFileName) {
  appLogFile.WriteLine("--- UpdateText START ---");
  iFileName = fso.getAbsolutePathName(iFileName);
  if (!fso.FileExists(iFileName)) {
    appLogFile.WriteLine("ファイルが存在しません=" + iFileName);
    return 9;
  }
  var oFileName = fso.BuildPath(fso.GetParentFolderName(iFileName), fso.GetBaseName(iFileName) + "_new." + fso.GetExtensionName(iFileName));
  withWriteFile(oFileName, function(oFile) {
    withEachLine(iFileName, function(line) {
      oFile.WriteLine("更新:" + line);
      WScript.Echo("更新:" + line);
    });
  });
  appLogFile.WriteLine("--- UpdateText  END  ---");
}

//-----------------------------------------------------------------------------
function CreateCsv(oFileName) {
  appLogFile.WriteLine("--- CreateCsv START ---");
  withWriteFile(oFileName,function (file) {
    var ary = new Array();
    for(var row=0;row<100;row++) {
      ary = [];
      for(var col=0;col<10;col++) {
        ary.push(row.toString() + "-" + col.toString());
      }
      file.WriteLine(ary.join(","));
    }
  });
  appLogFile.WriteLine("--- CreateCsv  END  ---");
}

//-----------------------------------------------------------------------------
function ReadCsv(iFileName) {
  appLogFile.WriteLine("--- ReadCsv START ---");
  if (!fso.FileExists(iFileName)) {
    appLogFile.WriteLine("ファイルが存在しません=" + iFileName);
    return 9;
  }
  withEachLine(iFileName,function (line) {
    var ary = line.split(",");
    var str = ary.join("/")
    WScript.Echo(ary.toString());
    WScript.Echo(str);
  });
  appLogFile.WriteLine("--- ReadCsv  END  ---");
}

//-----------------------------------------------------------------------------
function UpdateCsv(iFileName) {
  appLogFile.WriteLine("--- UpdateCsv START ---");
  iFileName = fso.getAbsolutePathName(iFileName);
  if (!fso.FileExists(iFileName)) {
    appLogFile.WriteLine("ファイルが存在しません=" + iFileName);
    return 9;
  }
  var oFileName = fso.BuildPath(fso.GetParentFolderName(iFileName), fso.GetBaseName(iFileName) + "_new." + fso.GetExtensionName(iFileName));
  withWriteFile(oFileName, function(oFile) {
    withEachLine(iFileName, function(line) {
      var ary = line.split(",");
      var str = ary.join("/");
      oFile.WriteLine(str);
      WScript.Echo(str);
    });
  });
  appLogFile.WriteLine("--- UpdateCsv  END  ---");
}

//-----------------------------------------------------------------------------
function CreateExcel(oFileName) {
  appLogFile.WriteLine("--- CreateExcel START ---");
  newWorkbook(fso.getAbsolutePathName(oFileName), false, function (workbook) {
    var worksheet =  workbook.Worksheets(1);
    worksheet.Cells(1,1).value = "No";
    worksheet.Cells(1,2).value = "DATA1";
    worksheet.Cells(1,3).value = "DATA2";
    for(var row=2;row<100;row++){
      worksheet.Cells(row,1).value = row - 1;
      worksheet.Cells(row,2).value = "D1" + "-" + (row-1);
      worksheet.Cells(row,3).value = "D2" + "-" + (row-1);
    }
  });
  appLogFile.WriteLine("--- CreateExcel  END  ---");
}

//-----------------------------------------------------------------------------
function ReadExcel(iFileName) {
  appLogFile.WriteLine("--- ReadExcel START ---");
  if (!fso.FileExists(iFileName)) {
    appLogFile.WriteLine("ファイルが存在しません=" + iFileName);
    return 9;
  }
  withWorkbook(fso.getAbsolutePathName(iFileName), false, true, function (workbook) {
    var worksheet = workbook.Worksheets(1);
    var row = 1;
    for (;;) {
      var n = worksheet.Cells(row, 1).value;
      if (!n || n == "") break;
      WScript.Echo(worksheet.Cells(row,1).value + worksheet.Cells(row,2).value + worksheet.Cells(row,3).value);
      row++;
    }
  });
  appLogFile.WriteLine("--- ReadExcel  END  ---");
}

//-----------------------------------------------------------------------------
function ReadExcelSheets(iFileName) {
  appLogFile.WriteLine("--- ReadExcelSheets START ---");
  if (!fso.FileExists(iFileName)) {
    appLogFile.WriteLine("ファイルが存在しません=" + iFileName);
    return 9;
  }
  withWorkbook(fso.getAbsolutePathName(iFileName), false, true, function (workbook) {
    var worksheetsCount = workbook.Worksheets.Count;
    for (var shtNo=1;shtNo<=worksheetsCount;shtNo++) {
      var worksheet = workbook.Worksheets(shtNo);
      WScript.Echo("sheet_name=" + worksheet.Name);
      var row = 1;
      for (;;) {
        var n = worksheet.Cells(row, 1).value;
        if (!n || n == "") break;
        WScript.Echo(worksheet.Cells(row, 1).value + worksheet.Cells(row, 2).value + worksheet.Cells(row, 3).value);
        row++;
      }
    }
  });
  appLogFile.WriteLine("--- ReadExcelSheets  END  ---");
}

//-----------------------------------------------------------------------------
function UpdateExcel(iFileName) {
  appLogFile.WriteLine("--- UpdateExcel START ---");
  if (!fso.FileExists(iFileName)) {
    appLogFile.WriteLine("ファイルが存在しません=" + iFileName);
    return 9;
  }
  withWorkbook(fso.getAbsolutePathName(iFileName), false, true, function (workbook) {
    var absFileName = fso.GetAbsolutePathName(iFileName)
    var worksheet = workbook.Worksheets(1);
    var row = 1;
    for (;;) {
      var n = worksheet.Cells(row, 1).value;
      if (!n || n == "") break;
      worksheet.Cells(row,4).value = worksheet.Cells(row,1).value + "(new)";
      row++;
    }
    var oFileName = fso.BuildPath(fso.GetParentFolderName(absFileName),  fso.GetBaseName(absFileName) + "_new." + fso.GetExtensionName(absFileName));
    workbook.SaveAs(fso.getAbsolutePathName(oFileName));
  });
  appLogFile.WriteLine("--- UpdateExcel  END  ---");
}

//-----------------------------------------------------------------------------
function CreateJson(oFileName) {
  appLogFile.WriteLine("--- CreateJson START ---");

  var outJson = {};
  var dmyAry = [1, 2, 3];
  outJson.a = "A";
  outJson.b = "B";
  outJson.nList = dmyAry;
  var jsonStr = JSON.stringify(outJson);
  withWriteFile(oFileName, function (file) {
    file.WriteLine(jsonStr);
  });

  appLogFile.WriteLine("--- CreateJson  END  ---");
}

//-----------------------------------------------------------------------------
function ReadJson(iFileName) {
  appLogFile.WriteLine("--- ReadJson START ---");
  if (!fso.FileExists(iFileName)) {
    appLogFile.WriteLine("ファイルが存在しません=" + iFileName);
    return 9;
  }
  withAllLine(iFileName, function (jsonStr) {
    var objJson = JSON.parse(jsonStr);
    WScript.Echo("objJson.a=" + objJson.a);
    WScript.Echo("objJson.b=" + objJson.b);
    WScript.Echo("objJson.nList=" + objJson.nList);
  });
  appLogFile.WriteLine("--- ReadJson  END  ---");
}
