// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const filebrowser = document.getElementById('uploadfile');
const msgDiv = document.getElementById('msginfo');
const openDirBtn = document.getElementById('opendir');
const fs = require('fs');
const ospath = require('path');
openDirBtn.style.display ='none';
//const filename = document.getElementById('outputfile').value;
const Excel = require('exceljs/modern.nodejs');
const outputDir = require('electron').remote.app.getPath('desktop');
const items_map = {};
const baseTime = new Date('1900-01-01 00:00:00');
openDirBtn.onclick=()=>{

  const dir =   ospath.resolve(outputDir,'Abson');
  const { shell } = require('electron')
  shell.openExternal(dir)
  
}

function makeCountry(str){
  for(var i=0;i<CountryArr.length;i++){
    console.log('makeCountry',CountryArr[i],str+'-',CountryArr[i].includes((str+'-')));
    if(CountryArr[i].includes(str+'-')){
      console.log('makeCountry done',CountryArr[i].split("-")[1])
      return CountryArr[i].split("-")[1];
    }
  }
  return str;
}

function makeDou(str) {
  if (str < 10) return `0${str}`;
  else return `${str}`;
}

filebrowser.onchange = (e) => {

  const file = e.target.files[0];
  const {
    path
  } = file;
  console.log('path', path)
  msgDiv.innerHTML='';
  var PizZip = require('pizzip');
  var Docxtemplater = require('docxtemplater');

  

  //Load the docx file as a binary
  var tempfilename = ospath.resolve(__dirname, 'template.docx')
  var content = fs
    .readFileSync(tempfilename, 'binary');

  var zip = new PizZip(content);

  var doc = new Docxtemplater();
  doc.loadZip(zip);

  var workbook = new Excel.Workbook();
  msgDiv.innerHTML +=`正在读取文件。。。<br>`
  workbook.xlsx.readFile(path)
    .then(function () {
      var worksheet = workbook.getWorksheet('Disp w.34');
      if(!worksheet){
        alert('excel里面缺少页面 Disp w.34')
        msgDiv.innerHTML =`excel里面缺少页面 Disp w.3<br>`
        return;
      }
      msgDiv.innerHTML +=`读取文件成功<br>`
      worksheet.eachRow(function (row, rowNumber) {
        if (rowNumber > 2) {
          const ddate = new Date(baseTime.getTime() + (row.getCell(17) - 1) * 24 * 3600 * 1000);
          if(!row.getCell(2).text)return;
          const item = {
            ponum: row.getCell(2).text.substr(2),
            ponumstr: row.getCell(2).text,
            order: row.getCell(1).text,
            jworder: row.getCell(3).text,
            itemno: row.getCell(5).text,
            day: makeDou(ddate.getDate()),
            month: makeDou(ddate.getMonth() + 1),
            year: makeDou(ddate.getFullYear()),
            monthstr: ddate.toDateString().split(" ")[1],
            dest: row.getCell(4).text,
            destZh:makeCountry(row.getCell(4).text.split(",")[1].replace(/ /g,'')),
            description: row.getCell(6).text,
            amount: row.getCell(7).text,
            price: row.getCell(15).text,
            money: (row.getCell(7).text * row.getCell(15).text).toFixed(2),
            nowyear:makeDou(new Date().getFullYear()),
            nowmonth:makeDou(new Date().getMonth() + 1),
            nowday:makeDou(new Date().getDate()),
          }
          if (items_map[row.getCell(2).text]) {
            items_map[item.ponumstr].totalmoney += (item.money*1);
            items_map[item.ponumstr].clients.push(item);
          } else {
            items_map[item.ponumstr] = Object.assign({},item);
            items_map[item.ponumstr].totalmoney = (item.money*1);
            items_map[item.ponumstr].clients = [item];
          }
         
        }
      });

      for (var key in items_map) {
        //set the templateVariables
        console.log('key',items_map[key])
        items_map[key].totalmoney = items_map[key].totalmoney.toFixed(2);
        doc.setData(items_map[key]);
        const filename = `${key}_${items_map[key].nowyear}_${items_map[key].nowmonth}_${items_map[key].nowday}.docx`

        try {
          // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
          doc.render()
        } catch (error) {
          var e = {
            message: error.message,
            name: error.name,
            stack: error.stack,
            properties: error.properties,
          }
          console.log(JSON.stringify({
            error: e
          }));
          // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
          throw error;
        }

        var buf = doc.getZip()
          .generate({
            type: 'nodebuffer'
          });

        const dir =   ospath.resolve(outputDir,'Abson');
        if (!fs.existsSync(dir)) {
          fs.mkdirSync(dir)
        }
        const outputFile = ospath.resolve(dir, filename);
        try{
          fs.writeFileSync(outputFile, buf);
          msgDiv.innerHTML +=`${filename} 已生成完毕<br>`
        }catch(e){
          msgDiv.innerHTML +=`${filename} 文件已打开,生成失败<br>`
        }
        
        
        
        //alert('完成')
        

      }
      msgDiv.innerHTML +=`全部生成完毕`;
      openDirBtn.style.display ='block';



      // use workbook

    });





  

}