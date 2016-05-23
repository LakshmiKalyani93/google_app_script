function intiateAlert() {
  try{
    alertAllocManager();
  }catch(e){
    sendReportInMail(e);
  }
}
function alertAllocManager(){
  
  var resSS = SpreadsheetApp.openById("1p-YaHi_6_uQ_CDr0wXNb6fIPqWpn5MAhUn0u52HkDyk");
  var resSheet = resSS.getSheetByName(buildSheetName(new Date().getMonth()));
  
  if(resSheet!= null){
    
    var indexArr = [];
    var managerData = [];
    var namesArr = [];
    var managerArr = [];
    var titleData = resSheet.getRange(2, 1, 1, resSheet.getLastColumn()).getValues();  
    var rangeCount = getBillabilityRangeCount(titleData);    
    Logger.log("RangeCount : "+rangeCount);
    
    if(rangeCount == 4){
      managerData = resSheet.getRange(3,23,resSheet.getLastRow(),1).getValues();
      indexArr=[7,11,15,19];
    }else{
      managerData = resSheet.getRange(3,27,resSheet.getLastRow(),1).getValues();
      indexArr=[7,11,15,19,23];
    }
    
    var uniqueManagerArr = removeDups(managerData);
    Logger.log("Unique Managers : "+ uniqueManagerArr);
        
    var lastRow = resSheet.getLastRow();
    for(i=3;i<lastRow;i++){
      for(j in indexArr){
        var project = resSheet.getRange(parseInt(i), parseInt(indexArr[j])).getValue();
        if(project.toString().equals("") && !(project.toString().indexOf('Resigned')>-1)){
          namesArr.push(resSheet.getRange(parseInt(i), 1).getValue());
          if(rangeCount == 4){
            managerArr.push(resSheet.getRange(parseInt(i), 23).getValue());
          }else{
            managerArr.push(resSheet.getRange(parseInt(i), 27).getValue()); 
          }
          break;
        }   
      }   
    }
    
    var resultArr = getUniqueResultantArray(uniqueManagerArr,managerArr,namesArr);
    var namesArr =  resSheet.getRange(3, 1, resSheet.getLastRow(), 1).getValues();
    var data='';
    for(j in resultArr){
      Logger.log("Resultant Array : "+ resultArr[j]);
      data = broadcastReportInMail(resultArr[j],resSheet,namesArr,data);
    }
    
    Logger.log("Data : ========="+data);
    var recipientsTO = "kalyani.kolimali@mtuity.com"+","+"vijay@mtuity.com";
    var message = "Hi Team, "
    var subject = "Testing - Allocation Details not filled yet – Alert!";
    
    var html = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'+
'<html xmlns="http://www.w3.org/1999/xhtml">'+
'<head>'+
'<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />'+
'<title>Resource Availability Status</title>'+
'</head>'+
''+
'<body>'+
'<table style="border:solid 1px #c7d7ff; font-family:Arial, Helvetica, sans-serif; font-size:13px;" align="center" border="0" width="620" cellpadding="0" cellspacing="0">'+
'  <tbody>'+
'    <tr>'+
'      <td style="padding:20px 35px;text-align:center;margin:0;background:#212a41" valign="top"><img src="http://staging.mtuity.com/devicetracking/logo.png" alt="Logo" style="text-align:center" height="72" width="165" class="CToWUd" /></td>'+
'    </tr>'+
'    <tr>'+
'      <td style="padding:0;margin:0" valign="top"><table align="center" border="0" width="620" cellpadding="0" cellspacing="0">'+
'          <tbody>'+
'            <tr>'+
'              <td><table width="620" border="0" cellspacing="0" cellpadding="0" style="margin-bottom:20px">'+
'                  <tr>'+
'                    <td align="center" style="background:#337ab7; padding:10px;"><span style="font-size:16px !important;color:#fff">Below are resources for whom allocation is yet to be filled for current month.<br />'+
'                      Please fill it at the earliest in the Resource Allocation Sheet</span></td>'+
'                  </tr>'+
'                  <tr>'+
'                    <td><table width="620" cellspacing="0" cellpadding="3">'+
'                        <tr>'+
'                          <td width="220" align="left" style="background:#444f5c;  border-left:solid 1px #333;border-bottom:solid 1px #333; color:#fff"><strong>Name</strong></td>'+
'                          <td width="187" align="left" style="background:#444f5c; border-left:solid 1px #333;border-bottom:solid 1px #333; color:#fff"><strong>Technology Group</strong></td>'+
'                          <td width="193" align="left" style="background:#444f5c;  border-left:solid 1px #333;border-bottom:solid 1px #333; color:#fff; border-right:solid 1px #333"><strong>Allocation Manager</strong></td>'+
'                        </tr>'+data+
'                      </table></td>'+
'                  </tr>'+
'                </table></td>'+
'            </tr>'+
'          </tbody>'+
'        </table></td>'+
'    </tr>'+
'    <tr>'+
'      <td style="padding:10px;margin:0;background-color:#444f5c;text-align:center;font-size:14px;color:#ffffff" valign="top" width="600">Mtuity Resource Allocation Team</td>'+
'    </tr>'+
'  </tbody>'+
'</table>'+
'</body>'+
'</html>';
     
    if(!data.toString().equals('')){
    MailApp.sendEmail(recipientsTO, subject, message, { htmlBody: html}); 
    }
  }      
}

function broadcastReportInMail(arr,sheet,namesArr,data){
  
  if(sheet!=null){
    var techGrpArr = [];
    var len = arr.length; 
    if(len > 0){
      for(i=1;i<len;i++){
        for(j in namesArr){
          if(arr[i].toString().equals(namesArr[j].toString())){ 
            Logger.log("Name and index " + arr[i] +"======"+(parseInt(j)+3));
            techGrpArr[0] = "TechGrp";
            techGrpArr[i] = sheet.getRange(parseInt(j)+3,2).getValue();   
          }  
        }
      }
      Logger.log("TechGrp : "+techGrpArr +" ===================== Lengths : "+ techGrpArr.length +" and " +arr.length);
    }
  }
  
  for(k=1;k<arr.length;k++){
  //data = arr[k]+" "+techGrpArr[k]+" "+arr[0]+";\n\n\n\n\n\n\n\n\n\n\n\n\n"+ data;
   data = '<tr><td align="left" style="border-left:solid 1px #333;border-bottom:solid 1px #333;">'+
     arr[k]+'</td><td align="left"  style="border-left:solid 1px #333;border-bottom:solid 1px #333;">'+
     techGrpArr[k]+'</td><td align="left"  style="border-left:solid 1px #333;border-bottom:solid 1px #333; border-right:solid 1px #333">'+
     arr[0]+'</td></tr>'+ data;
  }
  return data;
}

function getUniqueResultantArray(uniqueManagerArr,managerArr,namesArr){
  
  var arr =[];

  for(j in uniqueManagerArr){
    var tempArr =[];
    for(k in managerArr){  
      if(!uniqueManagerArr[j].toString().equals("") && uniqueManagerArr[j].toString().equals(managerArr[k].toString())){
        tempArr[0] = managerArr[k];
        tempArr.push(namesArr[k]);
      }
    }
    if(tempArr.length>0){
      arr.push(tempArr);
    }
  }
  
  return arr;
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
  for(i in array){      
    var data = array[i];
    var duplicate = false;    

    if(data.toString().equals("")){
      continue;
    }
    
    for(j in newArr){  
      if(newArr[j].toString().equals("")){
        continue;
      }
      if(data.join() == newArr[j].join()){
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
  Logger.log("SheetNameValue =========", sheetNameVal);
}

function sendReportInMail(e){
  
  var recipientsTO= "kalyani.kolimali@mtuity.com";
   //*****var recipientsTO = "kalyani.kolimali@mtuity.com"+","+"vijay@mtuity.com";
  var message = "Hi Vijay/Kalyani, "
  var subject = " Allocation Details not filled yet – Failure Report";
  var html = '<body>'
       + '<p>'+message+'</p>'
       + '<p> Allocation Details not filled yet – Failure Report Details :</p>'
       + '<p> Error Message : '+e.message+'.'+'</p>' 
       + '<p> FileName : '+e.fileName +'.'+'</p>'
       + '<p> Line Number : '+e.lineNumber +'.'+'</p>'
       +'<p> </p>'
       +'<p> </p>'
       +'<p>Thanks & Regards,<br></br> Mtuity ResAlloc Tracking Team. </p>'
       +'</body>'
       MailApp.sendEmail(recipientsTO, subject, message, { htmlBody: html});
  
}
