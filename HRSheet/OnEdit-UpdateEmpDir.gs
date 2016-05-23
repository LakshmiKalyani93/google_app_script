function intiateUpdate(){
  
 var sheet = SpreadsheetApp.getActive();
 ScriptApp.newTrigger("onSheetEdit")
   .forSpreadsheet(sheet)
   .onEdit()
   .create();
}

function onSheetEdit(){
  
  try{
    
  var hrSS = SpreadsheetApp.getActiveSpreadsheet();
  var hrSheet=hrSS.getActiveSheet();
  
  var activeCell = hrSheet.getActiveCell();
  var activeRowIndex = activeCell.getRow();
  var activeColIndex = activeCell.getColumn();
  var activeCellData = activeCell.getValue();
  
  var colTitle = hrSheet.getRange(1, activeColIndex).getValue();
  
  if(!colTitle.toString().equals("") &&( colTitle.toString().equals("Name") || colTitle.toString().equals("Designation"))){
    
    var emailId = hrSheet.getRange(activeRowIndex, 4).getValue();
    
    var empDirSheetId = "1Mw1Utgbqt43F7H4L9tALvO1e9zlCw1XP4MZ_-BlBvaI";
    var empSS = SpreadsheetApp.openById(empDirSheetId);
    var empSheet = empSS.getSheetByName("Mtuity");
    
    var empEmailsData = empSheet.getRange(2, 4, empSheet.getLastRow(), 1).getValues() ;
    
    
    for(i=0;i<empEmailsData.length;i++){
       
      if(emailId.toString().equals(empEmailsData[i].toString())){
                
        if(colTitle.toString().equals("Name")){       
          empSheet.getRange(parseInt(i)+2, 1).setValue(activeCellData);
        }else{
         empSheet.getRange(parseInt(i)+2, 9).setValue(activeCellData);
        }     
      }          
    }
    
  }
   
  }catch(e){
    sendReportInMail(e);
  }
   
}


function sendReportInMail(e){
  
   var recipientsTO = "vijay@mtuity.com";
   var message = "Hi Vijay, "
   var subject = "HR Sheet - Emp Directory Update Script Failure";
   var html = '<body>'
       + '<p>'+message+'</p>'
       + '<p> HR Sheet Failure Report Details :</p>'
       + '<p> Error Message : '+e.message+'.'+'</p>' 
       + '<p> FileName : '+e.fileName +'.'+'</p>'
       + '<p> Line Number : '+e.lineNumber +'.'+'</p>'
       +'<p> </p>'
       +'<p> </p>'
       +'<p>Thanks & Regards,<br></br> Mtuity Emp Directory Tracking Team </p>'
       +'</body>'
     MailApp.sendEmail(recipientsTO, subject, message, { htmlBody: html});
  
}
