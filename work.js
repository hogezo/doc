"use strict";
//const XLSX= require('xlsx');
const utils = XLSX.utils;

let workbook ="";
let worksheet ="";
let modelName = "";
let fileName = "";
var outputBuf = [];

//readFromFile();

var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
document.getElementById('btn').onclick = function(){
  var files = document.getElementById('files'), f = files.files[0];
  modelName = document.forms.id_form1.id_textBox1.value;
  fileName = escape(f.name);
  var timeStamp = f.lastModifiedDate.toLocaleDateString();
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = e.target.result;
    if(!rABS) data = new Uint8Array(data);
    workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});
    writeOutbuf("//modelName=" + modelName);
    writeOutbuf("//fileName=" + fileName);
    writeOutbuf("//settingFiletimeStamp=" + timeStamp);
    parseExcelFile();
    /* DO SOMETHING WITH workbook HERE */
  };
  if(rABS) reader.readAsBinaryString(f); else reader.readAsArrayBuffer(f);
}

function handleDrop(e) {
  e.stopPropagation(); e.preventDefault();
//  var files = e.dataTransfer.files, f = files[0];
  var files = e.target.files, f = files[0];
  //modelNameをテキストBOXから取得
  modelName = document.forms.id_form1.id_textBox1.value;
  fileName = escape(f.name);
  var timeStamp = f.lastModifiedDate.toLocaleDateString();
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = e.target.result;
    if(!rABS) data = new Uint8Array(data);
    workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});
    writeOutbuf("//modelName=" + modelName);
    writeOutbuf("//fileName=" + fileName);
    writeOutbuf("//settingFiletimeStamp=" + timeStamp);
    parseExcelFile();
    /* DO SOMETHING WITH workbook HERE */
  };
  if(rABS) reader.readAsBinaryString(f); else reader.readAsArrayBuffer(f);
}


function readFromFile(){
  // コマンドライン引数の取得
  const args = process.argv;
  const runtime = args.shift(), script = args.shift();
  modelName = args.shift();
  const file_name = args.shift();
  fileName = (file_name === undefined) ? "test.xlsx" : file_name;
  //const modelName = (model_name === undefined) ? ".c" : model_name;
  if (modelName === undefined) {
    console.log("[USAGE] compileSW.js (modelName) (filename)");
    process.exit();
  }
  console.log("//modelName=" + modelName);
  console.log("//fileName=" + fileName);

  workbook = XLSX.readFile(fileName, {
      cellDates: true
  });
  parseExcelFile();
}

function parseExcelFile(){
  //シートの読み込み
  let sheetnames = workbook.SheetNames;
  //let modelName = "METIS_MF3";
  //let worksheet = workbook.Sheets['コンパイルオプション_16S_17S_18S'];
  //console.log(sheetnames);
  let range ="";
  for(let i=0;i<sheetnames.length;i++){
    if(sheetnames[i].search(`16S`) > -1){
      worksheet = workbook.Sheets[sheetnames[i]];
      range = worksheet['!ref'];
      break;
    }
    else{
  //    console.log("skip "+sheetnames[i]);
    }
  }
  searchModel(worksheet, range);
}

function searchModel(workSheet, cellRange){
  let infCol = 0;
  let targetRow = 0;

  let rangeVal = utils.decode_range(cellRange);
//  console.log(utils.decode_range(cellRange));
  modelName = (modelName == "") ? "VESTA-MF2" : modelName;

  for (let r=0 ; r <= rangeVal.e.r ; r++) {
    let adr = utils.encode_cell({c:infCol, r:r});
    let cell = worksheet[adr];
    if(cell != undefined){
      if(cell.v == "model"){
//        console.log("Model found!! col="+targetRow);
        break;
      }
      else{
//        console.log("not found "+adr+" "+cell.v+" "+cell.t);
      }
    }
    targetRow++;
  }
  let c = 0;
  for (c=0 ; c <= rangeVal.e.c ; c++) {
    let adr = utils.encode_cell({c:c, r:targetRow});
    let cell = worksheet[adr];
    if(cell != undefined){
      if(cell.v == modelName){
//        console.log("Model found!! col="+c);
        break;
      }
      else{
//        console.log("not found "+adr+" "+cell.v+" "+cell.t);
      }
    }
  }
  dumpSwitch(c, rangeVal.e.r)
}

function dumpSwitch(tCol, eRow){
//  console.log("#if defined("+modelName+")");
  writeOutbuf("#if defined("+modelName+")");
  for (let r=0 ; r <= eRow ; r++) {
    let adr = utils.encode_cell({c:0, r:r});
    let cell = worksheet[adr];
    if(cell != undefined){
      if(cell.v == "1"){
        let compString = worksheet[utils.encode_cell({c:1, r:r})];
        let commentString = worksheet[utils.encode_cell({c:2, r:r})];
//        console.log(compString.v);
//        console.log(worksheet[utils.encode_cell({c:tCol, r:r})]);
        var isEnable = worksheet[utils.encode_cell({c:tCol, r:r})];
        if(isEnable != undefined){
          if("○" == isEnable.v){
//            console.log("#define "+compString.v+" "+"1");
            writeOutbuf("/*" + commentString.v + "*/");
            writeOutbuf("  #define "+compString.v+" "+"1");
          }
          else{
//            console.log("//"+compString.v);
            writeOutbuf("/*" + commentString.v + "*/");
            writeOutbuf("  #define "+compString.v+" "+"(0)");
          }
        }
      }
      else{
  //      console.log("not found "+adr+" "+cell.v+" "+cell.t);
      }
    }
  }
//  console.log("#endif");
  writeOutbuf("#endif");
  writeOutbuf("// " + modelName + "END");
  showOutputResult();
  alert("作成完了");
}

function writeOutbuf(outString){
//  outputBuf.push('<li>', outString, '</li>');
  outputBuf.push(outString, '<br>');
}

function showOutputResult(){
//  document.getElementById('list').innerHTML = '<ul>' + outputBuf.join('') + '</ul>';
  document.getElementById('list').innerHTML = '<h5>' + outputBuf.join('') + '</h5>';
}

function readCell(workSheet, cellRange){
  let rangeVal = utils.decode_range(cellRange);
  for (let r=rangeVal.s.r ; r <= rangeVal.e.r ; r++) {
      for (let c=rangeVal.s.c ; c <= rangeVal.e.c ; c++) {
          let adr = utils.encode_cell({c:c, r:r});
          let cell = worksheet[adr];
          if(cell != undefined){
            console.log(`${adr} type:${cell.t} value:${cell.v} text:${cell.w}`);
          }
      }
  }
}
