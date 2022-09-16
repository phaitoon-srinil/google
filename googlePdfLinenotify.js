//สำหรับฟอร์ม
function onFormSubmit() {  
  var token1 = ["LINE_TOKEN-1"];//โทเค่นไลน์ผู้มีสิทธิ์อนุมัติ
  
  var form = FormApp.openById('GOOGLE_FORM_ID'); //ใส่หมายเลข id ของ google form ที่ให้ผู้ใช้กรอกข้อมูล 
  var fRes = form.getResponses();
  var formResponse = fRes[fRes.length - 1];
  var itemResponses = formResponse.getItemResponses();
  
  var ss = SpreadsheetApp.openById('GOOGLE_SHEET_ID'); //ใส่หมายเลข id ของ google sheet ที่ใช้แสดงข้อมูลของ google form
  var sheet = ss.getSheetByName('RESPONSE_SHEET_NAME'); //ใส่ชื่อของ sheet ที่ใช้เก็บข้อมูล
  var row = sheet.getActiveRange().getLastRow()+1;

//ให้ไปสร้างลิสต์อนุมัติในชีตแผ่น 2 เพื่อดึงค่ามาแสดงที่ชีตแผ่น 1
   var dynamicList = ss.getSheetByName('APPROVED_SHEET_NAME').getRange('A1:A3');//ใส่ชื่อของ sheet ของตัวเลือกในการอนุมัติ 
                                                                                //และช่วงข้อมูล (A1:A3)
   var rangeRule = SpreadsheetApp.newDataValidation().requireValueInRange(dynamicList).build();
   sheet.getRange(row,6).setDataValidation(rangeRule); //คอลัมน์ที่ 6 คือ ช่องที่แสดงข้อความของผลการอนุมัติ

  var msg = itemResponses[0].getResponse() + ' : ส่งเรื่องขออนุมัติการเข้าอบรม' +'\n'+ ss.getUrl();;

  sendLineNotify(msg,token1);
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

//สำหรับชีต
function approve() {
   var token2 = "o1iE4Hm5FQXsrpw2WsUTCob91Fw4ib9lKbeHt6u42kF";//โทเค่นของกลุ่มไลน์ที่ต้องการให้ส่งข้อความไปแจ้งเตือน
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = ss.getActiveSheet();
   var row = sheet.getActiveRange().getRow();
   var cellvalue = sheet.getActiveCell().getValue().toString();
   var sheetName = sheet.getName();   
 
   var date = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/");
   var year = Number(Utilities.formatDate(new Date(), "GMT+7", "yyyy"));
   var thaiyear = Number(Utilities.formatDate(new Date(), "GMT+7", "yyyy"))+543;
   //var time = Utilities.formatDate(new Date(), "GMT+7", "HH:mm");
  
   var data1 = sheet.getRange(row, 3,row).getValue();//ชื่อ-นามสกุล
   var data2 = sheet.getRange(row, 4,row).getValue();//ชื่อหน่วยงาน
   var data3 = sheet.getRange(row, 5,row).getDisplayValue();//วันที่เข้าร่วมอบรม

   var message = 'แจ้งผลการสมัครเข้าอบรม: '+cellvalue+'\n'+'ชื่อ-สกุลผู้เข้าอบรม:'+data1+'\n'+'สถานที่อบรม:'+data2+'\n'+'วันที่เข้าอบรม:'+data3+'\n'+'วันที่อนุมัติ คือ:'+ date+year;

     if (cellvalue == 'อนุมัติ' || cellvalue == 'ไม่อนุมัติ' ) {
        sendLineNotify(message, token2);
        createBulkPDFs ();
      }
    }

function createBulkPDFs () {
  const pdfFolder = DriveApp.getFolderById("Pdf_FOLDER_ID");
  const tempFolder = DriveApp.getFolderById("Temp_FOLDER_ID");
  const docFile = DriveApp.getFileById("Doc_FOLDER_ID");
  
  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = currentSheet.getActiveCell().getRow();
  const data = currentSheet.getRange(row, 1,1,currentSheet.getLastColumn()).getValues()[0]; 
  //data[2] คือ ชื่อ-นามสกุล, data[3] คือ =ชื่อหน่วยงาน, data[4] คือ วันที่เข้าร่วมอบรม, data[5] คือ ผลการอนุมัติ
  createPDF(data[2], data[3],data[4], data[5], data[2], docFile, tempFolder, pdfFolder);

}

function createPDF (Name,Place,date,approve,pdfName,docFile,tempFolder,pdfFolder) {
  const tempFile = docFile.makeCopy(tempFolder); 
  const tempDocFile = DocumentApp.openById(tempFile.getId()); 
  const body = tempDocFile.getBody(); 
  
  body.replaceText("{ชื่อ-นามสกุล}", Name); 
  body.replaceText("{ชื่อหน่วยงาน}", Place); 
  body.replaceText("{วันที่เข้าร่วมอบรม}", date); 
  body.replaceText("{ผลการอนุมัติ}", approve);
  
  tempDocFile.saveAndClose(); 
  const pdfContentBlob = tempFile.getAs(MimeType.PDF); 
  const pdfFile = pdfFolder.createFile(pdfContentBlob).setName(pdfName);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.COMMENT);
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveRange().getLastRow();
  var email = sheet.getRange(row,2).getValue().toString(); //คอลัมน์ที่2 ใน google sheet คือ ข้อมูลที่อยู่อีเมล

  MailApp.sendEmail(email, 'การอนุมัติการไปอบรม', 'ดาวน์โหลดเอกสารได้ที่\n'+pdfFile.getUrl());
  tempFolder.removeFile(tempFile);

}

