function onOpen() {
var ss = SpreadsheetApp.getActive();
var items = [
    {name: 'View Contact Info', functionName: 'getContact'},
    {name: 'View Leave Plan', functionName: 'getLeavePlan'},
    {name: 'Create ResAlloc Template', functionName: 'createResAllocTemplate'}

  ];
  ss.addMenu('Paradigm', items);
}


function getLeavePlan(){
   Browser.msgBox("Coming Soon!");
}

function getContact() {
  
  var empMaster = SpreadsheetApp.openById("0Av4CaBAyrYmDdHZsa1NDTGhvV2JlQzA4T2c4eVhkWmc") 
  var empMasterSheet = empMaster.getSheetByName("PCS");
  var emplist = empMasterSheet.getDataRange().getValues();
  
  var row = SpreadsheetApp.getActiveRange().getRow();
  var contact = SpreadsheetApp.getActive().getDataRange().getValues();
  

  foundIndex = -1;
  for (var i = 1; i<emplist.length; i++) {
      if(emplist[i][0] == contact[row-1][0]){
        foundIndex = i;
        Browser.msgBox("Name: "+emplist[i][0]+"\\n"+
                       "Location: "+emplist[i][2]+"\\n"+
                       "Email: "+emplist[i][3]+"\\n"+
                       "Phone: "+emplist[i][4]+"\\n"+
                       "Skype: "+emplist[i][5]+"\\n");        
        break;
      }      
  }
  
  if(foundIndex == -1){
      Browser.msgBox("Sorry! Contact not found in Employee Directory");
  }   
}


function createResAllocTemplate(){
  
  try{
    prepareResAllocTemplate();
  }catch(e){
    sendReportInMail(e);
  }
 
}



function prepareResAllocTemplate() {
  
  var  result = Browser.inputBox('Enter month & year(Eg: 4,2016 ; use no spaces) to create Resource Allocation Template Sheet? ', Browser.Buttons.OK_CANCEL);
  if (result.equals("cancel")) {
        return;
  } 
  
  var arr = result.split(",");  
  if(parseInt(arr[0])>0 && parseInt(arr[0])<13){
      Browser.msgBox('Request is placed successfully.!!');
      processSheetCreation(arr);  
  } else{
    Browser.msgBox('Sorry ! Cannot process invalid data, please try again later !!');
  }
  
}


function processSheetCreation(arr){
  
  var monthVal = arr[0];
  var year = arr[1];
  
  var resourceSS = SpreadsheetApp.openById("1p-YaHi_6_uQ_CDr0wXNb6fIPqWpn5MAhUn0u52HkDyk"); 
  var templateSheet = resourceSS.getSheetByName("ResAllocTemplateSheet");   
  var buildSheetNameVal = buildSheetName(parseInt(monthVal)-1); 
  var noOfSheets = resourceSS.getSheets().length;   
  resourceSS.insertSheet(buildSheetNameVal, noOfSheets, {template : templateSheet });
  var month = parseInt(monthVal)-1; 
  formatTemplateSheet(year,month,resourceSS);
}



function formatTemplateSheet(year,month,ss){
  
    var monTitleRange = getWorkingWeekRange(year,month);
    var sheet = ss.getSheetByName(buildSheetName(month));
    var range = sheet.getRange(2, 1, 1, sheet.getLastColumn()); 
    var data = range.getValues();
    var count = findLastProjectColIndex(data); 
    Logger.log("Last Project Column Index Count : "+count); 
    if(count == 0 ){
     // Do nothing...
      return;
     }else{
       if(monTitleRange.length == 5){
         sheet.getRange(1, 7).setValue(monTitleRange[0]).setVerticalAlignment("middle").setHorizontalAlignment("center");
         sheet.getRange(1, 11).setValue(monTitleRange[1]).setVerticalAlignment("middle").setHorizontalAlignment("center");
         sheet.getRange(1, 15).setValue(monTitleRange[2]).setVerticalAlignment("middle").setHorizontalAlignment("center");
         sheet.getRange(1, 19).setValue(monTitleRange[3]).setVerticalAlignment("middle").setHorizontalAlignment("center");
         sheet.getRange(1, 23).setValue(monTitleRange[4]).setVerticalAlignment("middle").setHorizontalAlignment("center");       
       }else if(monTitleRange.length == 4){     
         Logger.log("Delete column module called...number of working days : "+monTitleRange);
        var tempVal = 0; 
        for(i=0;i<4;i++){
          var index = parseInt(count)+i;
          sheet.deleteColumn( index - tempVal);
          tempVal++;  
        }  
         sheet.getRange(1, 7).setValue(monTitleRange[0]).setVerticalAlignment("middle").setHorizontalAlignment("center");
         sheet.getRange(1, 11).setValue(monTitleRange[1]).setVerticalAlignment("middle").setHorizontalAlignment("center");
         sheet.getRange(1, 15).setValue(monTitleRange[2]).setVerticalAlignment("middle").setHorizontalAlignment("center");
         sheet.getRange(1, 19).setValue(monTitleRange[3]).setVerticalAlignment("middle").setHorizontalAlignment("center");
        
         for(i=3;i<sheet.getLastRow()+1;i++){
           
         var formula_24 = "=AVERAGE(H"+i+",L"+i+",P"+i+",T"+i+")";
         var formula_25 = "=AVERAGE(I"+i+",M"+i+",Q"+i+",U"+i+")";
         var formula_26 = "=AVERAGE(J"+i+",N"+i+",R"+i+",V"+i+")";
           
         sheet.getRange(i, 24).setFormula(formula_24);
         sheet.getRange(i, 25).setFormula(formula_25);
         sheet.getRange(i, 26).setFormula(formula_26);    
       }
         
       }else{
         return;
       }
  }
}

function findLastProjectColIndex(data){
   var count = 0;
   var colData = data[0];
   for (i=0;i<colData.length;i++){
     
    if(colData[i].toString().equals("Project")){  
      count++;
      }   
     
     if(count == 5){
        return parseInt(i)+1;
      }else{
        continue;
      }
    }
}


function getWorkingWeekRange(year,month){
  
  var startDate = new Date(year, month, 1);
  var endDate = new Date(year, month + 1, 0);
  
  var startDay = startDate.getDay();
  var endDay = endDate.getDay();
  
  var month = startDate.getMonth();  
  
  var dayArr=[];
  
  Logger.log("StartDay: "+startDay +" StartDate : "+startDate.getDate());
  
  var monRangeArr =[];
  
  if(startDay == 0){ 
    dayArr[0] = (startDate.getDate()+ 1)+":"+(startDate.getDate()+ 5) ;  
    monRangeArr = setMonthRange(1,dayArr,endDate);
  }else if(startDay == 1){
    dayArr[0] = (startDate.getDate()+ 0)+":"+(startDate.getDate()+ 4) ;  
    monRangeArr = setMonthRange(1,dayArr,endDate);
  }else if(startDay == 2){
    dayArr[0] = (startDate.getDate()+ 0)+":"+(startDate.getDate()+ 3) ;  
    dayArr[1] = (startDate.getDate()+ 6)+":"+(startDate.getDate()+10) ;  
    monRangeArr = setMonthRange(2,dayArr,endDate);
  }else if(startDay == 3){
    dayArr[0] = (startDate.getDate()+ 0)+":"+(startDate.getDate()+ 2) ;  
    dayArr[1] = (startDate.getDate()+ 5)+":"+(startDate.getDate()+ 9) ; 
    monRangeArr = setMonthRange(2,dayArr,endDate);
  }else if(startDay == 4){
    dayArr[0] = (startDate.getDate()+ 0)+":"+(startDate.getDate()+ 8) ;  
    dayArr[1] = (startDate.getDate()+ 11)+":"+(startDate.getDate()+ 15) ;  
    monRangeArr = setMonthRange(2,dayArr,endDate);
  }else if(startDay == 5){
    dayArr[0] = (startDate.getDate()+ 0)+":"+(startDate.getDate()+7) ;  
    dayArr[1] = (startDate.getDate()+10)+":"+(startDate.getDate()+ 14) ; 
    monRangeArr = setMonthRange(2,dayArr,endDate);
  }else if(startDay == 6){ 
    dayArr[0] = (startDate.getDate()+2)+":"+(startDate.getDate()+ 6) ;  
    monRangeArr = setMonthRange(1,dayArr,endDate);
  }
 
  var monTitlesArr = [];
  var monTabs =[ "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" ];
  
  Logger.log("Range array and length : "+monRangeArr +"====="+monRangeArr.length);
  if(monRangeArr.length > 0){
    for(i=0;i<monRangeArr.length;i++){     
      monTitlesArr[i] = monTabs[month]+monRangeArr[i].toString().replace(":","  to "+monTabs[month]);
    }
    Logger.log("monTitlesArr "+ monTitlesArr);
  }
  return monTitlesArr;
  
}


function setMonthRange(strtIndx,dayArr,endDate){
  
  var endDt = endDate.getDate();

  var splitArr = [{}];
 
  for(n=strtIndx;n<5;n++){
    
    if(dayArr.length>0){
    splitArr = dayArr[n-1].toString().split(":");  
    var monday = parseInt(splitArr[0])+parseInt(7);
    var friday = parseInt(splitArr[1])+parseInt(7)
    if( monday <=endDt){
      
      if(friday <=endDt){    
        dayArr[n] = monday+":"+friday;        
      }else{    
        dayArr[n] = monday+":"+endDt;        
      }
      
    } 
   }
  }
     
  var datesArr =[];
  var c = 0;
  for(j=0;j<dayArr.length;j++){
     
     var tempArr = dayArr[j].toString().split(":");
     c=j+j;
     datesArr[c]= tempArr[0];
     datesArr[c+1]= tempArr[1];
         
    }
  Logger.log("Dates Array : "+datesArr);
  
  var tempStr = [];
  if(datesArr.length > 1){
  
  var comp_1 = datesArr[1] - datesArr[0];

  if(datesArr.length == 8){
    if(comp_1 > 3){
      
      tempStr[0] = datesArr[0]+":"+datesArr[1];
      tempStr[1] = datesArr[2]+":"+datesArr[3]; 
      tempStr[2] = datesArr[4]+":"+datesArr[5]; 
      tempStr[3] = datesArr[6]+":"+datesArr[7]; 
    }
      
  }else if(datesArr.length == 10){
    
    var comp_2 = datesArr[9]-datesArr[8];
    
    if(comp_1 < 2){
      
      tempStr[0] = datesArr[0]+":"+datesArr[3]; 
      tempStr[1] = datesArr[4]+":"+datesArr[5]; 
      tempStr[2] = datesArr[6]+":"+datesArr[7]; 
      tempStr[3] = datesArr[8]+":"+datesArr[9]; 
      
    }else if(comp_1 >= 2 && comp_1 <= 4){
    
      tempStr[0] = datesArr[0]+":"+datesArr[1]; 
      tempStr[1] = datesArr[2]+":"+datesArr[3]; 
      tempStr[2] = datesArr[4]+":"+datesArr[5]; 
     
      if(comp_2 < 2){  
        tempStr[3] = datesArr[6]+":"+datesArr[9]; 
      }else{  
        tempStr[3] = datesArr[6]+":"+datesArr[7]; 
        tempStr[4] = datesArr[8]+":"+datesArr[9]; 
      }
    }
   }
  }
  
 Logger.log("Formatted Array : "+tempStr);
  
  return tempStr;
}

  
function buildSheetName(monthVal){
  
  var currentYear = (new Date()).getYear();
    
  var monthTabs = [ "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December" ];
  
  Logger.log("Month : "+monthVal+"======= Year : "+currentYear+"======= SheetName : "+monthTabs[monthVal]+" "+currentYear);
 
  var sheetNameVal = monthTabs[monthVal]+" "+currentYear;
  
  return sheetNameVal;
  
  Logger.log("SheetNameValue =========", sheetNameVal);
}



function sendReportInMail(e){
  
  var recipientsTO= "kalyani.kolimali@mtuity.com";
   //*****var recipientsTO = "kalyani.kolimali@mtuity.com"+","+"vijay@mtuity.com";
   var message = "Hi Vijay/Kalyani, "
   var subject = "EMP Directory Script Failure Report";
   var html = '<body>'
       + '<p>'+message+'</p>'
       + '<p> EMP Directory Failure Report Details :</p>'
       + '<p> Error Message : '+e.message+'.'+'</p>' 
       + '<p> FileName : '+e.fileName +'.'+'</p>'
       + '<p> Line Number : '+e.lineNumber +'.'+'</p>'
       +'<p> </p>'
       +'<p> </p>'
       +'<p>Thanks & Regards,<br></br> Mtuity Emp Directory Tracking Team. </p>'
       +'</body>'
     MailApp.sendEmail(recipientsTO, subject, message, { htmlBody: html});
  
}
