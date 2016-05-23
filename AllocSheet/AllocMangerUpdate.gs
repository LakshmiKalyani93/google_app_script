function initiatePMUpdate() {
  
 var sheet = SpreadsheetApp.getActive();
 ScriptApp.newTrigger("onPMEdit")
   .forSpreadsheet(sheet)
   .onEdit()
   .create();
}

function onPMEdit(){
  
  try{
  
  var activeSS = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet=activeSS.getActiveSheet();
  
  var activeCell = activeSheet.getActiveCell();
  var activeRowIndex = activeCell.getRow();
  var activeColIndex = activeCell.getColumn();
  var activeCellData = activeCell.getValue();
    
  var activePMTitle = activeSheet.getRange(2,activeColIndex).getValue();
  var activeResName = activeSheet.getRange(activeRowIndex, 1).getValue();
    
  var resSS = SpreadsheetApp.openById("1p-YaHi_6_uQ_CDr0wXNb6fIPqWpn5MAhUn0u52HkDyk");
  var currentMonth = (new Date()).getMonth(); 
  var resSheet = resSS.getSheetByName(buildSheetName(parseInt(currentMonth))) 
    
    
  if(resSheet!=null && resSheet.getSheetName().equals(activeSheet.getSheetName()) && (activePMTitle.toString().equals("Allocation Manager") || activePMTitle.toString().equals("Technology Group")||activePMTitle.toString().equals("Primary Skills")||activePMTitle.toString().equals("Secondary Skills"))){
         
    var templateSheet = resSS.getSheetByName("ResAllocTemplateSheet");
    
    // update template sheet...
    updatePMModule(templateSheet,activeResName,activeCellData,5,activePMTitle);
    
    //update next month sheet...
    updateFutureSheetsIfExists(1,resSS,activeResName,activeCellData,activePMTitle);
    
    //update next next month sheet....
    updateFutureSheetsIfExists(2,resSS,activeResName,activeCellData,activePMTitle);
    
    //update next next next month sheet....
    updateFutureSheetsIfExists(3,resSS,activeResName,activeCellData,activePMTitle);

   }
    
  }catch(e){
    sendReportInMail(e);  
  }
 
}

function updateFutureSheetsIfExists(sheetNo,resSS,activeResName,activeCellData,activePMTitle){
  
  var currentMonth = (new Date()).getMonth();  
  var monSheet = resSS.getSheetByName(buildSheetName(parseInt(currentMonth)+parseInt(sheetNo)));
  
  if(monSheet!=null){
  
  var billabilityColData = monSheet.getRange(2, 1, 1, monSheet.getLastColumn()).getValues();
  var billabilityRange = getBillabilityRange(billabilityColData);
  
  Logger.log("Sheet & Billability Range : " + monSheet + " & "+ billabilityRange );

  updatePMModule(monSheet,activeResName,activeCellData,billabilityRange,activePMTitle);

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


function updatePMModule(sheet,activeResName,activeCellData,range,activePMTitle){
  
  if(sheet!=null){
    
   var namesData = sheet.getRange(3, 1, sheet.getLastRow(), 1).getValues();
      
    for(i=0;i<namesData.length;i++){
        
      if(activeResName.toString()!=null && activeResName.toString().equals(namesData[i].toString())){
        
        var index = parseInt(i)+3;
        
        Logger.log("Index : "+ index);
        if(activePMTitle.toString().equals("Allocation Manager")){
          if( parseInt(range) == 5){ 
            sheet.getRange(index, 27).setValue(activeCellData).setVerticalAlignment("bottom").setHorizontalAlignment("left").setFontFamily("Times New Roman"); 
          }else{
            sheet.getRange(index, 23).setValue(activeCellData).setVerticalAlignment("bottom").setHorizontalAlignment("left").setFontFamily("Times New Roman"); 
          }
        }else if(activePMTitle.toString().equals("Technology Group")){
          sheet.getRange(index, 2).setValue(activeCellData).setVerticalAlignment("bottom").setHorizontalAlignment("left").setFontFamily("Times New Roman"); 
        }else if(activePMTitle.toString().equals("Primary Skills")){
          sheet.getRange(index, 3).setValue(activeCellData).setVerticalAlignment("bottom").setHorizontalAlignment("left").setFontFamily("Times New Roman");          
        }else if(activePMTitle.toString().equals("Secondary Skills")){
          sheet.getRange(index, 4).setValue(activeCellData).setVerticalAlignment("bottom").setHorizontalAlignment("left").setFontFamily("Times New Roman");           
        }
      }    
    }
    
  }
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


function sendReportInMail(e){
  
  var recipientsTO="kalyani.kolimali@mtuity.com"
   //var recipientsTO = "kalyani.kolimali@mtuity.com"+","+"vijay@mtuity.com";
   var message = "Hi Vijay,"
   var subject = "ResAlloc - PM Update Script Failure";
   var html = '<body>'
       + '<p>'+message+'</p>'
       + '<p> ResAlloc PM Update Failure Report: </p>'
       + '<p> Error Message : '+e.message+'.'+'</p>' 
       + '<p> FileName : '+e.fileName +'.'+'</p>'
       + '<p> Line Number : '+e.lineNumber +'.'+'</p>'
       +'<p> </p>'
       +'<p> </p>'
       +'<p>Thanks & Regards,<br></br> Mtuity Emp Directory Tracking Team </p>'
       +'</body>'
     MailApp.sendEmail(recipientsTO, subject, message, { htmlBody: html});
  
}
