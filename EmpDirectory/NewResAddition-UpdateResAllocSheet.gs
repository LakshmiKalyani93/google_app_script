function syncResAllocEmpDirectorySheets(){
  try{
    processSyncResAllocSheet();  
  }catch(e){
    sendReportInMail(e);
  } 
}

// module that updates the sheet contents
function processSyncResAllocSheet(){
    
  var sourceSheetId = "1Mw1Utgbqt43F7H4L9tALvO1e9zlCw1XP4MZ_-BlBvaI";
  var sourceSS = SpreadsheetApp.openById(sourceSheetId);
  var sourceSheet=sourceSS.getSheetByName("Mtuity");
  
  var empNamesListInDirectory = sourceSheet.getRange(2,1,sourceSheet.getLastRow(),1).getValues();
  
  
  var allocSheetId = "1p-YaHi_6_uQ_CDr0wXNb6fIPqWpn5MAhUn0u52HkDyk";
  var allocSS = SpreadsheetApp.openById(allocSheetId);
    
  var allocTemplateSheet = allocSS.getSheetByName("ResAllocTemplateSheet");
  var empNamesListInResAlloc = allocTemplateSheet.getRange(3,1,allocTemplateSheet.getLastRow(),1).getValues();
  
  var newResIndexArray = findNewResourceIndex(empNamesListInDirectory,empNamesListInResAlloc);

  Logger.log("NewResourceIndexArray : "+newResIndexArray);
  
  // Resource allocation template sheet updation
  updateResAllocSheet(newResIndexArray,allocTemplateSheet,sourceSheet,5);
  
  // Updating the current month sheet
  updateFutureSheetsIfExists(0,allocSS,newResIndexArray,sourceSheet);
  
  // Updating the next month sheet
  updateFutureSheetsIfExists(1,allocSS,newResIndexArray,sourceSheet);
  
  // Updating the next next month sheet
  updateFutureSheetsIfExists(2,allocSS,newResIndexArray,sourceSheet);

  // Updating the next next next month sheet
  updateFutureSheetsIfExists(3,allocSS,newResIndexArray,sourceSheet);
  
}


function updateFutureSheetsIfExists(sheetNo,allocSS,newResIndexArray,sourceSheet){
  
  var currentMonth = (new Date()).getMonth();  
  var monSheet = allocSS.getSheetByName(buildSheetName(parseInt(currentMonth)+parseInt(sheetNo)));
  
  if(monSheet!=null){
  
  var billabilityColData = monSheet.getRange(2, 1, 1, monSheet.getLastColumn()).getValues();
  var billabilityRange = getBillabilityRange(billabilityColData);
  Logger.log("Sheet & Billability Range : " + monSheet + " & "+ billabilityRange );
  updateResAllocSheet(newResIndexArray,monSheet,sourceSheet,billabilityRange); 
    
  }
}


// Module that returns the last project title column index...
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


function updateResAllocSheet(newResIndexArray,allocSheet,sourceSheet,billabilityRange){
  
  if(newResIndexArray.length > 0 && allocSheet!= null && sourceSheet!=null){
    
    var currentMonth = (new Date()).getMonth();  
    var breakPt = "";
    var currentMonAllocSheet = allocSheet.getSheetName().toString().equals(buildSheetName(currentMonth));
    
    if(currentMonAllocSheet){
      
      var index = [];
      if(billabilityRange == 4){
        index=[7,11,15,19];
      }else if(billabilityRange == 5){
        index=[7,11,15,19,23];
      }   
      breakPt = findJoinedBreaker(billabilityRange,allocSheet,index);
    }
    
    for(i in newResIndexArray){
   
      var resAllocLastRow = allocSheet.getLastRow()+1;
      
      var formula_10 = "=Max(1-SUM(H"+resAllocLastRow+":I"+resAllocLastRow+"),0)";
      var formula_14 = "=Max(1-SUM(L"+resAllocLastRow+":M"+resAllocLastRow+"),0)";
      var formula_18 = "=Max(1-SUM(P"+resAllocLastRow+":Q"+resAllocLastRow+"),0)";
      var formula_22 = "=Max(1-SUM(T"+resAllocLastRow+":U"+resAllocLastRow+"),0)";
      var light_gray_background = "#d9d9d9";
      var yellow_background = "#FFFF00";

      var value = newResIndexArray[i];
      
      allocSheet.getRange(resAllocLastRow, 1).setValue(sourceSheet.getRange(value, 1).getValue()).setFontFamily("Times New Roman");
      allocSheet.getRange(resAllocLastRow, 2).setValue(sourceSheet.getRange(value, 2).getValue()).setVerticalAlignment("bottom").setHorizontalAlignment("left");
      allocSheet.getRange(resAllocLastRow, 3).setValue(sourceSheet.getRange(value, 2).getValue()).setVerticalAlignment("bottom").setHorizontalAlignment("left");
      allocSheet.getRange(resAllocLastRow, 6).setValue(sourceSheet.getRange(value, 3).getValue()).setVerticalAlignment("bottom").setHorizontalAlignment("left");
      allocSheet.getRange(resAllocLastRow, 5).setValue(sourceSheet.getRange(value, 9).getValue()).setVerticalAlignment("bottom").setHorizontalAlignment("left"); 
      
      if(currentMonAllocSheet && (breakPt == 11 || breakPt == 15 || breakPt == 19 || breakPt == 23)){
        
        if(breakPt == 11){
          updateNotYetJoinedRange(7,allocSheet,resAllocLastRow);
          updateAvaialbiltiyRange(11,allocSheet,formula_14,yellow_background,resAllocLastRow);
          updateAvaialbiltiyRange(15,allocSheet,formula_18,yellow_background,resAllocLastRow);
          updateAvaialbiltiyRange(19,allocSheet,formula_22,yellow_background,resAllocLastRow); 
          updateLast3ColRange(billabilityRange,allocSheet,resAllocLastRow,light_gray_background,yellow_background); 
        }else if(breakPt == 15){   
          updateNotYetJoinedRange(7,allocSheet,resAllocLastRow);
          updateNotYetJoinedRange(11,allocSheet,resAllocLastRow);
          updateAvaialbiltiyRange(15,allocSheet,formula_18,yellow_background,resAllocLastRow);
          updateAvaialbiltiyRange(19,allocSheet,formula_22,yellow_background,resAllocLastRow); 
          updateLast3ColRange(billabilityRange,allocSheet,resAllocLastRow,light_gray_background,yellow_background); 
        }else if(breakPt == 19){
          updateNotYetJoinedRange(7,allocSheet,resAllocLastRow);
          updateNotYetJoinedRange(11,allocSheet,resAllocLastRow);
          updateNotYetJoinedRange(15,allocSheet,resAllocLastRow);
          updateAvaialbiltiyRange(19,allocSheet,formula_22,yellow_background,resAllocLastRow); 
          updateLast3ColRange(billabilityRange,allocSheet,resAllocLastRow,light_gray_background,yellow_background);
        }else if(breakPt == 23){
          updateNotYetJoinedRange(7,allocSheet,resAllocLastRow);
          updateNotYetJoinedRange(11,allocSheet,resAllocLastRow);
          updateNotYetJoinedRange(15,allocSheet,resAllocLastRow);          
          updateNotYetJoinedRange(19,allocSheet,resAllocLastRow);          
          updateLast3ColRange(billabilityRange,allocSheet,resAllocLastRow,light_gray_background,yellow_background);  
        }
        
      }else{
        
        updateAvaialbiltiyRange(7,allocSheet,formula_10,yellow_background,resAllocLastRow);
        updateAvaialbiltiyRange(11,allocSheet,formula_14,yellow_background,resAllocLastRow);
        updateAvaialbiltiyRange(15,allocSheet,formula_18,yellow_background,resAllocLastRow);
        updateAvaialbiltiyRange(19,allocSheet,formula_22,yellow_background,resAllocLastRow);
        updateLast3ColRange(billabilityRange,allocSheet,resAllocLastRow,light_gray_background,yellow_background); 
        
      }
    
    }  
    
  }
  
}

function updateLast3ColRange(billabilityRange,allocSheet,resAllocLastRow,light_gray_background,yellow_background){
  
        if(billabilityRange == 4){
        
          var formula_24 = "=AVERAGE(H"+resAllocLastRow+",L"+resAllocLastRow+",P"+resAllocLastRow+",T"+resAllocLastRow+")";
          var formula_25 = "=AVERAGE(I"+resAllocLastRow+",M"+resAllocLastRow+",Q"+resAllocLastRow+",U"+resAllocLastRow+")";
          var formula_26 = "=AVERAGE(J"+resAllocLastRow+",N"+resAllocLastRow+",R"+resAllocLastRow+",V"+resAllocLastRow+")";
          
          if("Hyd".equals(allocSheet.getRange(resAllocLastRow, 6).getValue().toString())){
             allocSheet.getRange(resAllocLastRow, 23).setValue("Padmaja");
          }else{
            allocSheet.getRange(resAllocLastRow, 23).setValue("Vijay");
          }
          allocSheet.getRange(resAllocLastRow, 24).setFormula(formula_24).setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground(light_gray_background);
          allocSheet.getRange(resAllocLastRow, 25).setFormula(formula_25).setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground(light_gray_background);
          allocSheet.getRange(resAllocLastRow, 26).setFormula(formula_26).setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground(light_gray_background);
      
      }else{
       
        var formula_26 = "=Max(1-SUM(X"+resAllocLastRow+":Y"+resAllocLastRow+"),0)";
        var formula_28 = "=AVERAGE(H"+resAllocLastRow+",L"+resAllocLastRow+",P"+resAllocLastRow+",T"+resAllocLastRow+",X"+resAllocLastRow+")";
        var formula_29 = "=AVERAGE(I"+resAllocLastRow+",M"+resAllocLastRow+",Q"+resAllocLastRow+",U"+resAllocLastRow+",Y"+resAllocLastRow+")";
        var formula_30 = "=AVERAGE(J"+resAllocLastRow+",N"+resAllocLastRow+",R"+resAllocLastRow+",V"+resAllocLastRow+",Z"+resAllocLastRow+")";
       
        updateAvaialbiltiyRange(23,allocSheet,formula_26,yellow_background,resAllocLastRow);
        if("Hyd".equals(allocSheet.getRange(resAllocLastRow, 6).getValue().toString())){
          allocSheet.getRange(resAllocLastRow, 27).setValue("Padmaja");
        }else{
          allocSheet.getRange(resAllocLastRow, 27).setValue("Vijay");
        } 
        allocSheet.getRange(resAllocLastRow, 28).setFormula(formula_28).setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground(light_gray_background);
        allocSheet.getRange(resAllocLastRow, 29).setFormula(formula_29).setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground(light_gray_background);
        allocSheet.getRange(resAllocLastRow, 30).setFormula(formula_30).setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground(light_gray_background);
      
      }
  
}

function updateNotYetJoinedRange(index,allocSheet,resAllocLastRow){
  
  allocSheet.getRange(resAllocLastRow, parseInt(index)).setValue("Not yet joined").setVerticalAlignment("middle").setHorizontalAlignment("center");
  allocSheet.getRange(resAllocLastRow, parseInt(index)+1).setValue(parseInt(0)).setVerticalAlignment("middle").setHorizontalAlignment("center");       
  allocSheet.getRange(resAllocLastRow, parseInt(index)+2).setValue(parseInt(0)).setVerticalAlignment("middle").setHorizontalAlignment("center"); 
  allocSheet.getRange(resAllocLastRow, parseInt(index)+3).setFormula(null).setValue(parseInt(0)).setVerticalAlignment("middle").setHorizontalAlignment("center");
}

function updateAvaialbiltiyRange(index,allocSheet,formula,yellow_background,resAllocLastRow){
  
  var currentMonAllocSheet = allocSheet.getSheetName().toString().equals(buildSheetName(new Date().getMonth()));
  if(currentMonAllocSheet){
    allocSheet.getRange(resAllocLastRow, parseInt(index)).setValue("Available").setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground(yellow_background);
    allocSheet.getRange(resAllocLastRow, parseInt(index)+1).setValue(parseInt(0)).setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground(yellow_background);       
    allocSheet.getRange(resAllocLastRow, parseInt(index)+2).setValue(parseInt(0)).setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground(yellow_background); 
    allocSheet.getRange(resAllocLastRow, parseInt(index)+3).setFormula(formula).setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground(yellow_background);
  }else{
    allocSheet.getRange(resAllocLastRow, parseInt(index)+1).setValue(parseInt(0)).setVerticalAlignment("middle").setHorizontalAlignment("center");       
    allocSheet.getRange(resAllocLastRow, parseInt(index)+2).setValue(parseInt(0)).setVerticalAlignment("middle").setHorizontalAlignment("center"); 
    allocSheet.getRange(resAllocLastRow, parseInt(index)+3).setFormula(formula).setVerticalAlignment("middle").setHorizontalAlignment("center");
  }
  
}

function findJoinedBreaker(count,sheet,index){
 
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



function findNewResourceIndex(directoryList,resAllocList){

  var isKeyFound = false;
  var tempNewResIndexArray =[];

  for(i in directoryList){
      
    if("Sridhar Gadhi".equals(directoryList[i].toString()) || "Suresh Chaparala".equals(directoryList[i].toString())){
        continue;
     }
   
    for(j in resAllocList){
        
      if(!directoryList[i].toString().equals("") && directoryList[i].toString().equals(resAllocList[j].toString())){
        isKeyFound = true;
        break;
       }
        
    }  
    
    if(!directoryList[i].toString().equals("")){
      if(isKeyFound){
        isKeyFound = false;
        continue;
      }else{
        tempNewResIndexArray.push(parseInt(i)+2);
      }   
    }
  }
  Logger.log("NewResIndexArray : " + tempNewResIndexArray);
  
  return tempNewResIndexArray;
}




function sendReportInMail(e){
  
   //var recipientsTO = "kalyani.kolimali@mtuity.com"+","+"vijay@mtuity.com";
  
   var recipientsTO = "kalyanikolimali93@gmail.com";
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
       +'<p>Thanks & Regards,<br></br> Mtuity Emp Directory Tracking Team </p>'
       +'</body>'
     MailApp.sendEmail(recipientsTO, subject, message, { htmlBody: html});
  
}

function buildSheetName(monthVal){
  
  var currentYear = (new Date()).getYear();
    
  var monthTabs = [ "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December" ];
  
  Logger.log("Month : "+monthVal+"======== Year : "+currentYear+"======= SheetName : "+monthTabs[monthVal]+" "+currentYear);
  var sheetNameVal = monthTabs[monthVal]+" "+currentYear;
  
  return sheetNameVal;
  
  Logger.log("SheetNameValue =========", sheetNameVal);
}

