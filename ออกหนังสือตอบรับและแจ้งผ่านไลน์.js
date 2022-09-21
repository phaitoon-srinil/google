//Credit: https://sysadmin.psu.ac.th/2021/05/03/
//Library PdfService: 1iePjnglUzelAuJJb-QykRcUUWYBSKiNGUWVljnNe03G9zWzSUGIRWLXa

function onFormSubmit() {  
  var token1 = ["LINE_TOKEN"];//โทเค่นไลน์ผู้มีสิทธิ์อนุมัติ
  
  var form = FormApp.openById('GOOGLE_FORM_ID'); //ใส่หมายเลข id ของ google form ที่ให้ผู้ใช้กรอกข้อมูล 
  var fRes = form.getResponses();
  var formResponse = fRes[fRes.length - 1];
  var itemResponses = formResponse.getItemResponses();
  
  var ss = SpreadsheetApp.openById('GOOGLE_SHEET_ID'); //ใส่หมายเลข id ของ google sheet ที่ใช้แสดงข้อมูลของ google form
  var sheet = ss.getSheetByName('การตอบแบบฟอร์ม 1'); //ใส่ชื่อของ sheet ที่ใช้เก็บข้อมูล
  var row = sheet.getActiveRange().getLastRow()+1;

  var msg = itemResponses[0].getResponse() + ' : ส่งเรื่องตอบรับเข้าร่วมงาน' +'\n'+ ss.getUrl();

  sendLineNotify(msg,token1);
  runPDF();
}

function sendLineNotify(message,token) {
  var options = {
    "method": "post",
    "payload": "message=" + message,
    "headers": {
    "Authorization": "Bearer " + token }
  };

UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}

function runPDF() {
  let sheetId = 'GOOGLE_SHEET_ID'; //ให้ใส่ id ของ google sheet ที่เก็บข้อมูลจาก google form
  let templateFileId = 'GOOGLE_SLIDE_ID'; //ให้ใส่ id ของ google slide ที่ใช้เป็นไฟล์ต้นแบบของหนังสือตอบรับ
  // let pdfFolder = DriveApp.getFoldersByName('Pdf').next()
  let pdfFolder = DriveApp.getFolderById("PDF_FOLDER_ID"); //ให้ใส่ Folder ID ที่จะให้เก็บไฟล์ pdf
  let templateFile = DriveApp.getFileById(templateFileId)
  let data = PdfService.initData(sheetId,'การตอบแบบฟอร์ม 1'); //ให้ใส่ชื่อ sheet ที่เก็บข้อมูลที่ได้จาก google form
  let option = {
    pdfFolder: pdfFolder,
    templateFile: templateFile,
    data: data,
    // image_column: ['รูปที่1','รูปที่2','รูปที่3'],
    // fileName: ['ชื่อหน้า','ชื่อกลาง','ชื่อหลัง'],
    image_column: ['IMAGE_FIELD_NAME'], //ให้ใส่ชื่อฟิลด์ที่ต้องกับรูปภาพที่ต้องการแทนที่ด้วยภาพจาก google form เช่่น {ลายเซ็นต์}
    fileName: ['FIELD_NAME'] //ชื่อของไฟล์ pdf ที่จะสร้างขึ้นมาจะเป็นชื่อเดียวกับข้อมูลใน field_name ที่กำหนด เช่น {ชื่อ-นามสกุล}
  }
    PdfService.createPDFFromSlide(option)
}
