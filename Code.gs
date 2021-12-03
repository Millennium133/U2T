//---------------------
//test
function test1(){
   let ss = SpreadsheetApp.getActiveSpreadsheet();
   let sheet = ss.getSheets()[1];

  // //let range = sheet.getRange(startRow);
  // //Logger.log(range.getRowIndex());
  // return sheet.getLastRow();
  let a = 'B2041';
  let range = sheet.getRange(a);
  // Logger.log(range.setBackgroundColor("#ea9999"));//red
  // Logger.log(range.setBackgroundColor("#b6d7a8"));//green
  // Logger.log(range.setBackgroundColor("#ffd966"));//yellow
  Logger.log(range.getValue());
  
  const check = range.getValue();
  // if (check.indexOf("\n")){
  //   Logger.log(check.indexOf("\n"));
  // }else{
  //   Logger.log(false);

  // }    
  sheet.getRange(2044,4).setValue("ผ่าน");
  sumHour = sheet.getRange(2044,6).getValue();
  command = `=if(${sumHour}>=20,"ผ่าน","ไม่ผ่าน")`
  sheet.getRange(2044,5).setValue(command);
  let fruits = [];
  a = "จิตวิทยาข้ามวัฒนธรรมในที่ทำงาน (8 ชั่วโมงการเรียนรู้) \nd| alihjflaih";
  
}
//---------------------



//---------------------
//custom menu
function initMenu(){
  const ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu("Aong's test function");
  menu.addItem("Main","main");
 // menu.addItem("test","test1");
  menu.addToUi();

}

function onOpen(){
  initMenu();
}
function extractOnlyNumber(input){
  let result = "";
  for (let i = 0;i < input.length;i++){
    if (!isNaN(input[i]))
      result += input[i];
  }
  return result;
}
//-------------------





//------------------
//main function --> change startRow and endRow here
function main(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[1];
  const startRow = "B1971";
  const endRow = "B2039";
  const validatedStartRow = parseInt(extractOnlyNumber(startRow),10);
  const validatedEndRow = parseInt(extractOnlyNumber(endRow),10);
  const range = validatedEndRow - validatedStartRow + 1;
  for(let row=0; row < range; row++){
      const category = sheet.getRange(validatedStartRow+row,8).getValue();
      doLiteracy(validatedStartRow+row, category);
      console.log(validatedStartRow+row);
  }
}
//----------------

function doLiteracy(curRow, category){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[1];

    const {url,columnCriterion,columnStartLit,columnEndLit,numerator} = config(category);

    let curColumn = sheet.getRange('A1').getColumn();
    let firstName,lastName,name,certificateCodes,sumHour = 0,arrayCerCode = [];
    let pass = 0;
    for(let column = 0; column < columnEndLit; column++){

      if (curColumn == 3){
        firstName = sheet.getRange(curRow,curColumn).getValue().trim();
      }
      else if (curColumn == 4){
        lastName = sheet.getRange(curRow,curColumn).getValue().trim();
        name = firstName + " " + lastName;    
      }
      if (curColumn >= columnStartLit && curColumn%2 == numerator){
        certificateCodes = (sheet.getRange(curRow,curColumn).getValue()).toString().trim();
        const certificateArr = certificateCodes.split(",");
        for(let num = 0; num < certificateArr.length; num++){
          validatedCer = checkCerCode(certificateArr[num].trim());
          let link = url + validatedCer;
          if (validatedCer != ""){
            let statusCode = getStatusCode(link)
            if (statusCode == 200){
              let {checkName, hour}= getNameAndHourThaiMooc(link);
              if (checkName == name && !arrayCerCode.includes(validatedCer)){
                arrayCerCode.push(validatedCer);
                console.log("typeof hour:",typeof hour);
                console.log("hour :",hour);
                sumHour += hour;
                console.log("SumHour :",sumHour);
                console.log("type of SumHour :",typeof sumHour);
                sheet.getRange(curRow,curColumn).setBackgroundColor('#b6d7a8');//green
              }else{
                pass++;
                sheet.getRange(curRow,curColumn).setBackgroundColor('#ea9999');//red
              }
            }else{//404 301 ...
              pass++;            
              sheet.getRange(curRow,curColumn).setBackgroundColor('#ffd966');//yellow
            }
          }else if (validatedCer == "" &&  sheet.getRange(curRow,curColumn+1).getValue() != ""){
            pass++;
            sheet.getRange(curRow,curColumn).setBackgroundColor('#ffd966');//yellow
          }
          if (pass > 0){
            break;
          }
        }
      }
      curColumn++;
    }
    console.log(sumHour);
    sheet.getRange(curRow,columnCriterion).setValue(sumHour);
    command = `=if(${sumHour}>=20,"ผ่าน","ไม่ผ่าน")`
    sheet.getRange(curRow,columnCriterion-1).setValue(command);
}



//sub function
function getNameAndHourThaiMooc(link){
  let statusCode = getStatusCode(link);
  if (statusCode == 200){
    let response = fetchData(link);
    let html = getHTML(response);
    let checkName = extractNameFromHtml(html,response.length);
    let hour = extractHourFromHtml(html,response.length);
    return {
      checkName:checkName,
      hour:hour
    };
  }
  else{
    return statusCode;
  }
}

function checkCerCode(certificateCode){
  let string = "";
  for (let i = 0;i<certificateCode.length;i++){
    if (certificateCode[i] != "|" || certificateCode[i] != "\n" ){
      string += certificateCode[i];
    }
  }
  return string;
}

function fetchData(url){
  let response = UrlFetchApp.fetch(url);
  return response;
}
function getHTML(response){
  let htmlText = response.getContentText();
  return htmlText;
}
function extractNameFromHtml(html,htmlLength){
  let balise = '<strong class="accomplishment-recipient hd-1 emphasized">'
  let cut = html.substring(html.indexOf( balise ), htmlLength);
  let value = cut.substring(balise.length, cut.indexOf("</strong>"));
  return value;
}
function extractHourFromHtml(html, htmlLength){
  let balise = '<span class="accomplishment-course-name">'
  let cut = html.substring(html.indexOf( balise ), htmlLength);
  let sentence = cut.substring(balise.length, cut.indexOf("</span>"));
  let hour = Number(sentence.substring(sentence.indexOf("(")+1, sentence.indexOf("ชั่วโมงการเรียนรู้")));
  return hour;
}

function getStatusCode(url) {
  let url_trimmed = url.trim();
  // Check if script cache has a cached status code for the given url
  let cache = CacheService.getScriptCache();
  let result = cache.get(url_trimmed);
  
  // If value is not in cache/or cache is expired fetch a new request to the url
  if (!result) {

    let options = {
      'muteHttpExceptions': true,
      'followRedirects': false
    };
    let response = UrlFetchApp.fetch(url_trimmed, options);
    let responseCode = response.getResponseCode();

    // Store the response code for the url in script cache for subsequent retrievals
    cache.put(url_trimmed, responseCode, 21600); // cache maximum storage duration is 6 hours
    result = responseCode;
  }

  return result;
}
//------------------------


//config
function config(literacy){

  let columnCriterion,columnStartLit,columnEndLit,numerator;
  const url = "https://lms.thaimooc.org/certificates/";
  if(literacy == "Financial literacy"){
    columnCriterion = 10;
    columnStartLit = 0;
    columnEndLit = 0;
    numerator = 0;
  }else if(literacy == "Language literacy"){
    columnCriterion = 12;
    columnStartLit = 17;
    columnEndLit = 34;
    numerator = 1;
  }else if(literacy == "Social literacy"){
    columnCriterion = 14;
    columnStartLit = 60;
    columnEndLit = 67;
    numerator = 0;
  }else if(literacy == "Digital literacy"){
    columnCriterion = 16;
    columnStartLit = 36;
    columnEndLit = 59 ;
    numerator = 0;
  }
  return {
      columnCriterion : columnCriterion,
      columnStartLit : columnStartLit,
      columnEndLit : columnEndLit,
      numerator : numerator,
      url:url,
  };

}

//-----------------------