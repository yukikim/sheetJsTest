  var inputSheet = function () {
    var input = document.getElementById("input_sheet");
    parseSheet(input.files[0], "Ver3.09-b", function (result) {
      // 変換後の処理
      console.log("result", result[9]);
      document.getElementById("output").textContent = (JSON.stringify(result, null, 2));
      // $("pre#result").html(JSON.stringify(result, null, 2));
    });
  };

  // var target = document.getElementById("output");
// フォームで入力されたExcelのsheetNameシートをオブジェクトにする。
  var parseSheet = function (file, sheetName, callback) {
    var reader = new FileReader();
    reader.onload = function (e) {
      var unit8 = new Uint8Array(e.target.result);
      var workbook = XLSX.read(unit8, {type: "array"});
      // 通常は1行目がヘッダ行となる。{header: 1} を指定で配列形式となる。
      var sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {header: 1});
      // 基本的な使い方であれば、sheet_to_jsonの結果だけでOK
      // 下記は変換を行う例
      var result = [];
      // 0行目から末尾まで走査
      for (var i = 0; i < sheet.length; i++) {
        var row = sheet[i];
        // 何か変換処理があれば行う
        console.log(i, row);
        result.push(row);
      }
      callback(result);
    };
    reader.readAsArrayBuffer(file);
  };

