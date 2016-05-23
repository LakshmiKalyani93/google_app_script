var property = PropertiesService.getScriptProperties();

function deletionOfResourceModule(){
  
  var resourceSS = SpreadsheetApp.openById("1p-YaHi_6_uQ_CDr0wXNb6fIPqWpn5MAhUn0u52HkDyk");
  var currentMonth = (new Date()).getMonth();
  var buildSheetNameVal = buildSheetName(currentMonth);
  
  property.setProperty("CurntMonth", currentMonth);
  
  var resourceSheet = resourceSS.getSheetByName(buildSheetNameVal);   
  var resourceData = getData(resourceSS,resourceSheet,3);
   
  var mainSS = SpreadsheetApp.openById("1Mw1Utgbqt43F7H4L9tALvO1e9zlCw1XP4MZ_-BlBvaI");    
  var mainSheet = mainSS.getSheetByName("Mtuity");
  var mainData = getData(mainSS,mainSheet,3);
  
  var missedIndexArray = findDeletedRowIndex(resourceData,mainData,resourceSS,resourceSheet);

  Logger.log("Deleted Row for currently running sheet ===============", missedIndexArray);  
  resourceSS.setActiveSheet(resourceSheet);
  
  try{
  setResignedInCurrentMonthSheet(resourceSheet,missedIndexArray);
  //updateFutureSheetsData(resourceSS, mainData);  
  }catch(e){
   sendReportInMail(e); 
  }

}


function setResignedInCurrentMonthSheet(resourceSheet,missedIndexArray){
  
  var numberOfResourceSheetCols = resourceSheet.getLastColumn();
  var data= resourceSheet.getRange(2, 1, 1, resourceSheet.getLastColumn()).getValues();
  var count = getBillabilityRange(data);
  
  Logger.log("count & Array : "+count +"======"+missedIndexArray);
  
  var index = [];
  
  if(count == 4){
    index=[7,11,15,19];
  }else if(count == 5){
        index=[7,11,15,19,23];
  }
  
  var resignedPos = findResignedBreaker(count,resourceSheet,index);
  
  Logger.log("Resigned Position : "+ resignedPos);  
  
  
  for(j=0;j<missedIndexArray.length;j++){
    
    var isChecked = checkResignedStatusIfAny(resourceSheet,missedIndexArray[j],index);
    
    if(isChecked){
      continue;
    }else{

      if(resignedPos  == 7){
        setCol1Data(resourceSheet,missedIndexArray[j],count);
      }else if(resignedPos == 11){
        setCol2Data(resourceSheet,missedIndexArray[j],count);
      }else if(resignedPos == 15){
        setCol3Data(resourceSheet,missedIndexArray[j],count);
      }else if(resignedPos == 19){
        setCol4Data(resourceSheet,missedIndexArray[j],count);    
      }else if(count == 5 && resignedPos == 23){
        setResignedData(resourceSheet,missedIndexArray[j],23,24,25,26);
      }
    
    }
    
  }  
  
}

function checkResignedStatusIfAny(sheet,row,indices){
  
  var isChecked = false;

  for(i=0;i<indices.length;i++){
      
    var str = sheet.getRange(row,indices[i]).getValue();
    if(str.toString().indexOf('Resigned')>-1){
      isChecked = true;
      break;
    }else{
      continue;
    }
  }
  
  return isChecked;
}

function setCol1Data(sheet,row,count){
  
  setResignedData(sheet,row,7,8,9,10);
  setCol2Data(sheet,row,count);
}

function setCol2Data(sheet,row,count){
     setResignedData(sheet,row,11,12,13,14);
     setCol3Data(sheet,row,count);
}

function setCol3Data(sheet,row,count){
  
  setResignedData(sheet,row,15,16,17,18);
  setCol4Data(sheet,row,count);
}

function setCol4Data(sheet,row,count){
   setResignedData(sheet,row,19,20,21,22);    
    if(count == 5){
      setResignedData(sheet,row,23,24,25,26);
    }
}

function setResignedData(sheet,row, col1,col2,col3,col4){
    
    var light_gray_background = "#d9d9d9";
    var red_font_colr = "#FF0000"
    sheet.getRange(row, col1).setValue("Resigned").setFontFamily("Times New Roman").setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground(light_gray_background);
    sheet.getRange(row, col2).setValue("0").setFontColor(red_font_colr).setFontFamily("Times New Roman").setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground(light_gray_background);
    sheet.getRange(row, col3).setValue("0").setFontColor(red_font_colr).setFontFamily("Times New Roman").setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground(light_gray_background);
    sheet.getRange(row, col4).setFormula(null).setValue("0").setFontColor(red_font_colr).setFontFamily("Times New Roman").setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground(light_gray_background);
    
}

function findResignedBreaker(count,sheet,index){
 
   var monTabs =[ "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" ];
  
    var day = new Date().getDate();    
    var mon = new Date().getMonth();
  
  for(i=0;i<index.length;i++){
    
    var monTitle = sheet.getRange(1, index[i]).getValue().toString();
    var splitTo = monTitle.toString().split("  to "); 
    var splitMon = monTabs[mon];
    Logger.log("SplitTo : "+splitTo[0] +" and "+splitTo[1]);
    var comp1 = splitTo[0].toString().split(splitMon);
    var comp2 = splitTo[1].toString().split(splitMon);
    Logger.log("SplitTo : "+comp1[1] +" and "+comp2[1]);
  
  if(day >=comp1[1] && day <= comp2[1]){
    
    return index[i];
  } else{
    continue;
  }    
 
  }
    
}
function getBillabilityRange(data){
  
   var count = 0;
   var colData = data[0];
 
   for (i=0;i<colData.length;i++){
     
    if(colData[i].toString().equals("Project")){  
      count++;
      }   
    }
  return count;
}

function updateFutureSheetsData(resourceSS,mainData){
 
  var currentMon = property.getProperty("CurntMonth");
  
  var nextMonthSheet = resourceSS.getSheetByName(buildSheetName(parseInt(currentMon)+1));
  removeResourcesInFuturesheet(resourceSS,nextMonthSheet,mainData);
  
  var nextNextMonthSheet = resourceSS.getSheetByName(buildSheetName(parseInt(currentMon)+2));
  removeResourcesInFuturesheet(resourceSS,nextNextMonthSheet,mainData);
  
  var nextNextNextMonthSheet = resourceSS.getSheetByName(buildSheetName(parseInt(currentMon)+3));
  removeResourcesInFuturesheet(resourceSS,nextNextNextMonthSheet,mainData);

  var templateSheet = resourceSS.getSheetByName("ResAllocTemplateSheet");
  removeResourcesInFuturesheet(resourceSS,templateSheet,mainData);
 
  Logger.log("CurrentMonth: "+currentMon+" Next Consecutive Sheet Names :  "+nextMonthSheet+", "+nextNextMonthSheet+", "+nextNextNextMonthSheet); 
  
}

function removeResourcesInFuturesheet(resourceSS,sheet,mainData){
  if(sheet!=null){
  var data = getData(resourceSS,sheet,3);
  var  delResArray = findDeletedRowIndex(data,mainData,resourceSS,sheet);
  iterateOverFutureSheets(sheet,resourceSS,delResArray);
  }
}

function findDeletedRowIndex(resourceData,mainData,resourceSS,sheet){
  
  var IsKeyFound = false;
  var tempIndexArray =[];
  var tempResourceArray=[]; 
  var currentMonth = new Date().getMonth();
  var curntSheetName = buildSheetName(currentMonth);
  
  if(sheet!= null){
    var sheetName = sheet.getSheetName().toString();
  }
  
  try{
  for(i in resourceData){  
   var resourceRow = resourceData[i];  
   var resLen = resourceRow.length;  
 
    for (j in mainData){
      var mainRow = mainData[j];   
      if(resourceData[i][0].toString().equals("")){
       continue; 
      } 
      if(!resourceData[i][0].toString().equals("") && resourceData[i][0].toString().equals(mainData[j][0].toString())){       
        IsKeyFound=true;
        break;
      }
     }
     if(!resourceData[i][0].toString().equals("")){
      
    if(IsKeyFound){
      IsKeyFound=false;
      continue;
    }else{
      tempResourceArray.push(resourceData[i][0]);
      tempIndexArray.push(parseInt(i)+3);
      continue;  
    }
    }
   }
  }catch(e){  
    sendReportInMail(e); 
  } 
  return tempIndexArray;
}


function sendReportInMail(e){
  
    var recipientsTO= "kalyani.kolimali@mtuity.com";
   //var recipientsTO = "kalyani.kolimali@mtuity.com"+","+"vijay@mtuity.com";
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


function iterateOverFutureSheets(sheetname,resourceSS,delResArray){
  
    if(sheetname !=null){
    
    Logger.log("Parsing the values in sheet : "+sheetname.getSheetName());  
      
    var mDataRange = sheetname.getRange(3, 1, sheetname.getLastRow(), 1);
    var mData = mDataRange.getValues();
    resourceSS.setActiveSheet(sheetname);

    try{
    var tempVal = 0;
    for(i=0;i<delResArray.length;i++){
                
        if(!delResArray[i].toString().equals("") ){
          var formatInx = parseInt(delResArray[i]);
          Logger.log("Delete Module : "+formatInx)
          sheetname.deleteRow(formatInx-tempVal);
          tempVal = tempVal + 1;
         
        }
    }   
      
   }catch(e){
     sendReportInMail(e);      
   }
  } 
}

function getData(ss,sheet,startRow){
  
  var data = [];
  if(sheet!=null){
  var data = sheet.getRange(startRow, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  ss.setActiveSheet(sheet);
  }
  return data;
  
}


function buildSheetName(monthVal){
  
  var currentYear = (new Date()).getYear();
    
  var monthTabs = [ "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December" ];
   
  var sheetNameVal = monthTabs[monthVal]+" "+currentYear;
  
  return sheetNameVal;
  
  Logger.log("SheetNameValue =========", sheetNameVal);
}




