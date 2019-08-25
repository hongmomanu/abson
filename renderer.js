// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const filebrowser = document.getElementById('uploadfile');
const msgDiv = document.getElementById('msginfo');
const openDirBtn = document.getElementById('opendir');
const refreshBtn = document.getElementById('refresh');
const fs = require('fs');
const ospath = require('path');
openDirBtn.style.display = 'none';
const Excel = require('exceljs/modern.nodejs');
const outputDir = require('electron').remote.app.getPath('desktop');

const baseTime = new Date('1900-01-01 00:00:00');
openDirBtn.onclick = () => {
  const dir = ospath.resolve(outputDir, 'Abson');
  const {
    shell
  } = require('electron')
  shell.openExternal(dir)

}
refreshBtn.onclick=()=>{
  location.reload();
}

function makeCountry(str) {
  for (var i = 0; i < CountryArr.length; i++) {
    console.log('makeCountry', CountryArr[i], str + '-', CountryArr[i].includes((str + '-')));
    if (CountryArr[i].includes(str + '-')) {
      console.log('makeCountry done', CountryArr[i].split("-")[1])
      return CountryArr[i].split("-")[1];
    }
  }
  return str;
}
function getDateOfISOWeek(w, y) {
  var simple = new Date(y, 0, 1 + (w - 1) * 7);
  var dow = simple.getDay();
  var ISOweekStart = simple;
  if (dow <= 4)
      ISOweekStart.setDate(simple.getDate() - simple.getDay() + 1);
  else
      ISOweekStart.setDate(simple.getDate() + 8 - simple.getDay());
  return ISOweekStart;
}
function getCurrentWeek(){
  var now = new Date();
  var onejan = new Date(now.getFullYear(),0,1);
  return ( Math.ceil((((now - onejan) / 86400000) + onejan.getDay()+1)/7));
}

function makeDou(str) {
  if (str < 10) return `0${str}`;
  else return `${str}`;
}
filebrowser.onclick = (e) => {
  filebrowser.value = null;
}
window.onkeydown=(e)=>{
  console.log(e.code);
  if(e.code === 'F12'){
    require('electron').remote.getCurrentWindow().openDevTools({mode:'bottom'});
  }
}
filebrowser.onchange = (e) => {
  const items_map = {};
  const file = e.target.files[0];
  const {
    path
  } = file;
  console.log('path', path)
  msgDiv.innerHTML = '';
  var PizZip = require('pizzip');
  var Docxtemplater = require('docxtemplater');



  //Load the docx file as a binary
  var tempfilename = ospath.resolve(__dirname, 'template.docx')
  var content = fs
    .readFileSync(tempfilename, 'binary');

  var zip = new PizZip(content);

  var doc = new Docxtemplater();
  doc.setOptions({linebreaks:true});
  doc.loadZip(zip);

  var workbook = new Excel.Workbook();
  msgDiv.innerHTML += `<div class="suc">正在读取文件。。。</div>`
  workbook.xlsx.readFile(path)
    .then(function () {
      var worksheet = workbook.getWorksheet('Disp w.34');
      if (!worksheet) {
        alert('excel里面缺少页面 Disp w.34')
        msgDiv.innerHTML = `<div class="err">excel里面缺少页面 Disp w.3！</div>`
        return;
      }
      msgDiv.innerHTML += `<div class="suc">读取文件成功。</div>`
      worksheet.eachRow(function (row, rowNumber) {
        if (rowNumber > 2) {
          //const ddate = new Date(baseTime.getTime() + (row.getCell(17) - 1) * 24 * 3600 * 1000);
          const week = row.getCell(13);
          const currentWeek = getCurrentWeek();
          const currentYear = new Date().getFullYear()
          const weekYear = currentWeek>week?currentYear+1:currentYear;
          const ddate = getDateOfISOWeek(week,weekYear);
          if (!row.getCell(2).text) return;
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
            destZh: makeCountry(row.getCell(4).text.split(",")[1].replace(/ /g, '')),
            description: row.getCell(6).text,
            amount: row.getCell(7).text,
            price: row.getCell(15).text,
            money: (row.getCell(7).text * row.getCell(15).text).toFixed(2),
            nowyear: makeDou(new Date().getFullYear()),
            nowmonth: makeDou(new Date().getMonth() + 1),
            nowday: makeDou(new Date().getDate()),
          }
          item.during = Factory_Map[item.itemno][0];
          item.seller = Factory_Map[item.itemno][1];
          item.sellerzh = Factory_Map[item.itemno][2];
          item.address = Factory_Map[item.itemno][3];
          item.addresszh = Factory_Map[item.itemno][4];
          item.tel = Factory_Map[item.itemno][5];
          item.packing = Factory_Map[item.itemno][6];
          item.sellerAgency = Factory_Map[item.itemno][7]?`\n${Factory_Map[item.itemno][7]}\n`:'';
          if (items_map[row.getCell(2).text]) {
            items_map[item.ponumstr].totalmoney += (item.money * 1);
            items_map[item.ponumstr].clients.push(item);
          } else {
            items_map[item.ponumstr] = Object.assign({}, item);
            items_map[item.ponumstr].totalmoney = (item.money * 1);
            items_map[item.ponumstr].clients = [item];
          }

        }
      });

      for (var key in items_map) {
        //set the templateVariables
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

        const dir = ospath.resolve(outputDir, 'Abson');
        if (!fs.existsSync(dir)) {
          fs.mkdirSync(dir)
        }
        const outputFile = ospath.resolve(dir, filename);
        try {
          fs.writeFileSync(outputFile, buf);
          msgDiv.innerHTML += `<div class="suc">${filename} 已成功！</div>`
        } catch (e) {
          msgDiv.innerHTML += `<div class="err">${filename} 文件已打开,生成失败</div>`
        }



        //alert('完成')


      }
      msgDiv.innerHTML += `<div class="suc">全部结束</div>`;
      openDirBtn.style.display = 'block';


    });







}