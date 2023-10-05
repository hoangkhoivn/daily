var express = require('express');
var router = express.Router();
const fs = require('fs');
const xlsx = require('xlsx');
const { format, parseISO } = require('date-fns');
var bodyParser = require('body-parser');
var urlencodedParser = bodyParser.urlencoded({ extended: false });
/* GET home page. */
var data = "Express";
var _weekday = "";
router.get('/', function (req, res, next) {
  res.render('index', { data_excel: data });
});
function ProcessExcel(res,mon,path_excel_file) {
  // Đường dẫn tới file Excel

  const filePath = 'nghia-2023.xlsx';

  // Đọc file Excel
  const workbook = xlsx.readFile(filePath);

  // Lấy danh sách các sheet trong file Excel
  const sheetNames = workbook.SheetNames;

  // Lấy dữ liệu từ sheet đầu tiên
  const firstSheetName = sheetNames[Number(mon)-1];
  const worksheet = workbook.Sheets[firstSheetName];

  // Chuyển đổi dữ liệu từ sheet thành JSON
  const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
  const dataArray = Object.values(jsonData);
  var num = 0;
  const data_html = [];
  for (i = 0; i < dataArray.length; i++) {
    var str = dataArray[i][16];
    if (str) {
      if (str == "■" || str == "○") {
        var h1 = dataArray[i][18];
        var date = dataArray[i - 7][1];
        const d = new Date(date);
        date2ms(d);
        //console.log(_weekday);
        let da = new Date(Math.round((d - 25569) * 864e5));
        const mDate = format(da, 'dd/MM/yyyy');
        //console.log(mDate);
        const products = [];
        for (ii = 1; ii < 7; ii++) {
          var product = dataArray[i - ii][3];
          if (product && product != "Product Name") {
            //console.log(product);
            products.push(product);
          }
        }
        var set_color;
        if (h1 < 1) {
          set_color = 'style="color: red;background-color:#000;"';
        } else {
          set_color = 'style="color: black; font-weight: normal;"';
        }
        if (str == "■" ) 
        {
          if (h1 > 1) {
            set_color = 'style="color: red;background-color:#000;"';
          }
        }
        if (str == "○" ) 
        {
          if (h1 < 7.75) {
            set_color = 'style="color: red;background-color:#000;"';
          }
        }

        if (products.length < 1) {
          products.push("-");
        }
        var html =
          `<li ${set_color} >
      <ul class="sub-ul">
        <li>
          <ul>
            <li>${_weekday}</li>
            <li><b>${mDate} </b></li>
          </ul>
        </li>
        <li>
          <ul>
            <li>${products.join(', ')}</li>
            <li> Giờ Làm: <b> ${h1} </b></li>
          </ul>

        </li>
      </ul>
  </li>`;
        data_html.push(html);
      }
    }
  }

  res.send(data_html.join(' '));
}
function date2ms(d) {
  let date = new Date(Math.round((d - 25569) * 864e5));
  const { getDay } = require('date-fns');
  const weekday = getDay(date);
  dayFromNumber(weekday);
  const mDate = format(date, 'dd/MM/yyyy');
  //console.log(mDate);
  return date;
}
function dayFromNumber(weekday) {
  var elements = [
    'Chủ nhật',
    'Thứ hai',
    'Thứ ba',
    'Thứ tư',
    'Thứ năm',
    'Thứ sáu',
    'Thứ bảy'];
  _weekday = elements[weekday];
  return elements[weekday];
}
router.post('/', function (req, res, next) {
  console.log("Mon :" + req.body.get_mon);
  console.log("Mon :" + req.body.get_path);
  ProcessExcel(res,req.body.get_mo,req.body.get_path);
}
);
module.exports = router;
