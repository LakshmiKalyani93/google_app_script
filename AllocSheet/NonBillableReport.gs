function initiateNonBillableReport() {
  try{
    getNonBillableProjReport();
  }catch(e){
    sendReportInMail(e);
  }
}

function getNonBillableProjReport(){
  
  var resSS = SpreadsheetApp.openById("1p-YaHi_6_uQ_CDr0wXNb6fIPqWpn5MAhUn0u52HkDyk");
  var resSheet = resSS.getSheetByName(buildSheetName(new Date().getMonth()));
   
  if(resSheet!= null){
    
    var indexArr = [];
    var nonBillIndicesArr = [];
    var nonBillProjArr = [];
    var projValArr = [];
    var titleData = resSheet.getRange(2, 1, 1, resSheet.getLastColumn()).getValues();  
    var rangeCount = getBillabilityRangeCount(titleData);
    
    Logger.log("RangeCount : "+rangeCount);
    
    if(rangeCount == 4){
      indexArr=[7,11,15,19];
      nonBillIndicesArr=[9,13,17,21];
    }else{
      indexArr=[7,11,15,19,23];
      nonBillIndicesArr=[9,13,17,21,25];
    }
    
    var lastRow = resSheet.getLastRow();
    
    for(i=3;i<=lastRow;i++){
      for(j in nonBillIndicesArr){
        var nonBillVal = resSheet.getRange(i, nonBillIndicesArr[j]).getValue();
        if(nonBillVal!=0){
          nonBillProjArr.push(resSheet.getRange(i,indexArr[j]).getValue().toString());
          projValArr.push([resSheet.getRange(i,indexArr[j]).getValue(),(nonBillVal/rangeCount)]);
        }
      }
    }
    
    var uniqueNonBillProj =  removeDups(nonBillProjArr);
    
    Logger.log("projValArr: "+projValArr.length+"========="+projValArr);  
    Logger.log("UninqueNonBillProj and length : "+uniqueNonBillProj.length+"========"+ uniqueNonBillProj);
    
    for(i in projValArr){
      Logger.log("ProjValArr : " + projValArr[i]);    
    }
    
    var resCount = 0;
  
//    for(i in uniqueNonBillProj){
//      for(j in projValArr){ 
//       // Logger.log("proj and projvalproj "+uniqueNonBillProj[i] +" "+projValArr[j][0]);
//        if(projValArr[j][0].toString().indexOf(uniqueNonBillProj[i])){
//          resCount = projValArr[j][1] + resCount;
//        }
//      }
//      Logger.log("UniqueProj and ResCount : "+uniqueNonBillProj[i] +" and "+ resCount);
//    }    
  }
}



function getBillabilityRangeCount(data){
  
  var count = 0;
  var colData = data[0];
 
  for (i=0;i<colData.length;i++){   
    if(colData[i].toString().equals("Project")){  
      count++;
    }   
  }
  return count;
}

function removeDups(array) {
 
  var newArr = []; 
  var splitArr =[];
  
  for(i in array){      
    var data = array[i];
    var duplicate = false; 
   
    for(j in newArr){  
      
      if(data.toString().indexOf("/") > -1){
        Logger.log("Reached data : "+ data);
        splitArr = data.toString().split("/");
      }
      
        for(k in splitArr){
          Logger.log("Split Value : "+ splitArr[k]);
          if(!splitArr[k].toString().equals("Available")){
            array.push(splitArr[k]); 
          }
        }
      
      if(!data.toString().equals("") && data.toString().equals(newArr[j].toString())){
        duplicate = true;
      }      
    }    
    if(!duplicate){
      newArr.push(data);
    }     
  }
  
  return newArr;
}

function buildSheetName(monthVal){
  
  var currentYear = (new Date()).getYear();    
  var monthTabs = [ "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December" ];
  var sheetNameVal = monthTabs[monthVal]+" "+currentYear;  
  return sheetNameVal;
}

function sendReportInMail(e){
  
  var recipientsTO= "kalyani.kolimali@mtuity.com";
   //*****var recipientsTO = "kalyani.kolimali@mtuity.com"+","+"vijay@mtuity.com";
  var message = "Hi Vijay/Kalyani, "
  var subject = " Non Billable Project – Failure Report";
  var html = '<body>'
       + '<p>'+message+'</p>'
       + '<p> Non Billable Project – Failure Report Details: </p>'
       + '<p> Error Message : '+e.message+'.'+'</p>' 
       + '<p> FileName : '+e.fileName +'.'+'</p>'
       + '<p> Line Number : '+e.lineNumber +'.'+'</p>'
       +'<p> </p>'
       +'<p> </p>'
       +'<p>Thanks & Regards,<br></br>Mtuity ResAlloc Tracking Team.</p>'
       +'</body>'
       MailApp.sendEmail(recipientsTO, subject, message, { htmlBody: html});
}
