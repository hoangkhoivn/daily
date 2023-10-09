var express = require('express');
var router = express.Router();
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { format, parseISO } = require('date-fns');
var bodyParser = require('body-parser');
var urlencodedParser = bodyParser.urlencoded({ extended: false });
/* GET home page. */
var data = "Express";
var _weekday = "";
var check_mon="";
router.get('/', function (req, res, next) {
  res.render('./index', { data_excel: data, check_file: "" });
});
router.get('/show', function (req, res, next) {
  res.render('./daily/showdaily', { mon: check_mon});
});
function ProcessExcel(res, json_data) {
  const dataArray =Object.values(json_data);
  const data_html = [];
  console.log("JsonData :" + dataArray.length);
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
  //console.log(data_html.join(' '));
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
  console.log("Get Mon Data111");
  ProcessExcel(res,req.body.json_data);
}
);
const multer = require('multer');
const storage = multer.memoryStorage(); // Lưu trữ tệp trong bộ nhớ
const upload = multer({ storage: storage });

// Đường dẫn để tải lên tệp Excel và xử lý
router.post('/upload', upload.single('excelFile'), (req, res) => {
  console.log(req.query.mon);
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }
  // Lấy dữ liệu từ tệp Excel (dữ liệu nằm trong req.file.buffer)
  const excelData = req.file.buffer;
// Sử dụng thư viện xlsx để đọc dữ liệu từ tệp Excel
const workbook = xlsx.read(excelData, { type: 'buffer' });
const sheetName = workbook.SheetNames[Number(req.query.mon)-1]; // Giả sử bạn muốn đọc từ sheet đầu tiên
  if (!sheetName) {
    return res.status(400).send('No sheet found in the Excel file.');
  }
  const sheet = workbook.Sheets[sheetName];
  const jsonData = xlsx.utils.sheet_to_json(sheet , { header: 1 });
  ProcessExcel(res, jsonData)
//res.send(JSON.stringify(data_json.join(' ')));

});
module.exports = router;
