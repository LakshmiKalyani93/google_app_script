function intiateResAvailability(){ 
  try{
    resAvailabiltiyModule();
  }catch(e){
    sendReportInMail(e);
  }
}

function resAvailabiltiyModule() {
  
  var resSS = SpreadsheetApp.openById("1p-YaHi_6_uQ_CDr0wXNb6fIPqWpn5MAhUn0u52HkDyk");
  var resSheet = resSS.getSheetByName(buildSheetName(new Date().getMonth()));
  
  if(resSheet!=null){
    
    var availResArr = [];
    var availTechGrpArr = [];
    var partialAvailResArr =[];
    var partialAvailTechGrpArr = [];
    var indexArr = [];
    var colIndex = "";
    
    var titleData = resSheet.getRange(2, 1, 1, resSheet.getLastColumn()).getValues();  
    var rangeCount = getBillabilityRangeCount(titleData);
    
    Logger.log("RangeCount : "+rangeCount);
    
    if(rangeCount == 4){
      indexArr=[7,11,15,19];
      colIndex= parseInt(26);
    }else{
      indexArr=[7,11,15,19,23];
      colIndex= parseInt(30);      
    }
    
    
    var colIndex = resSheet.getLastColumn();
    Logger.log("LastColIndex : "+ colIndex);
    
    var monthlyAvailColData = resSheet.getRange(3,colIndex,resSheet.getLastRow(),1).getValues();
    var techGrpData = resSheet.getRange(3,2,resSheet.getLastRow(),1).getValues();
    var uniqueTechGrp = removeDups(techGrpData);
   
    Logger.log("Unique Group : "+ uniqueTechGrp);
    Logger.log("Data : "+monthlyAvailColData);

    
    for(i in monthlyAvailColData ){
     
        if(monthlyAvailColData[i] == 1){
        
          availResArr.push(resSheet.getRange(parseInt(i)+3, 1).getValue());
          availTechGrpArr.push(resSheet.getRange(parseInt(i)+3, 2).getValue());
        
        }else if(monthlyAvailColData[i] != 0){
                   
          var isResigned = false;
          for(j in indexArr){
      
          var str = resSheet.getRange(parseInt(i)+3,indexArr[j]).getValue();
            if(str.toString().indexOf('Resigned')>-1){
              Logger.log("resigned called at row pos ......"+(parseInt(i)+3));
              isResigned = true;
              break;
            }
          }

          if(isResigned){
            continue; 
          }else{
            partialAvailResArr.push(resSheet.getRange(parseInt(i)+3, 1).getValue());
            partialAvailTechGrpArr.push(resSheet.getRange(parseInt(i)+3, 2).getValue()); 
          }
        }

     }
              
    var arr = getUniqueResultantArray(uniqueTechGrp,availTechGrpArr,availResArr);
    var partArr = getUniqueResultantArray(uniqueTechGrp,partialAvailTechGrpArr,partialAvailResArr);
    
    for(x in arr){
      Logger.log("Avail array : "+arr[x]);
    }
    
    for(y in partArr){
      Logger.log("Part array : "+partArr[y]); 
    } 
  
 } 
 simulateData(arr,partArr,resSheet,colIndex);
}


function simulateData(availArr, partArr,sheet,colIndex){
  
  for(i in availArr){  
    Logger.log("Available resources : "+availArr[i][0]);
    for(j in partArr){  
      Logger.log("Partially Available resources : "+partArr[j][0]);  
      if((!availArr[i][0].toString().equals("Done") || !partArr[j][0].toString().equals("Done")) 
        && availArr[i][0].toString().equals(partArr[j][0].toString())){
          var stream = availArr[i][0];
          var tempAvailArr = availArr[i];
          var tempPartAvailArr = partArr[j];
          availArr[i][0] = "Done";
          partArr[j][0] = "Done";
          broadcastEmailReport(stream,tempAvailArr,tempPartAvailArr,sheet,colIndex);
          break;
      }    
    }    
  }
  
   for(x in availArr){
     Logger.log("Res Avail array : "+availArr[x]);
     if(!availArr[x][0].toString().equals("Done")){
       var tempPartArr=[];
       broadcastEmailReport(availArr[x][0],availArr[x],tempPartArr,sheet,colIndex);       
     }  
   }
    
  for(y in partArr){
    Logger.log(" Res Part array : "+partArr[y]); 
     if(!partArr[y][0].toString().equals("Done")){
       var tempArr=[];
       broadcastEmailReport(partArr[y][0],tempArr,partArr[y],sheet,colIndex);       
     }
  } 
  
}


function broadcastEmailReport(stream,arr,partArr,sheet,colIndex){
  
  //var recipientsTO = "kalyani.kolimali@mtuity.com"+","+"vijay@mtuity.com";
  var recipientsTO = "kalyani.kolimali@mtuity.com";
  var message = "Hi Team, "
  var subject = "Testing - "+stream+" Resource Availability Status";
  var resData = "";
  var partResData="";
  
  if(sheet != null){
   
    var availLocArr = [];
    var partAvailLocArr = [];
    var partAvailVal =[];  
    var namesArr =  sheet.getRange(3, 1, sheet.getLastRow(), 1).getValues();
    var availRow ='';

    if(arr.length > 0){
      for(i=1 ;i < arr.length ;i++){         
        for(j in namesArr){    
          if(arr[i].toString().equals(namesArr[j].toString())){ 
            Logger.log("Name and index " + arr[i] +"======"+(parseInt(j)+3));
            availLocArr[0] = "Loc";
            availLocArr[i] = sheet.getRange(parseInt(j)+3,6).getValue();   
          }
        } 
      }
      for(x=1;x< arr.length; x++){
        resData = arr[x] +" "+ availLocArr[x]  + "; "+ resData; 
        availRow = '<tr><td align="left" style="border-left:solid 1px #333;border-bottom:solid 1px #333;">'+
          arr[x]+'</td><td align="left"  style="border-left:solid 1px #333;border-bottom:solid 1px #333;border-right:solid 1px #333;">'+
          availLocArr[x]+'</td></tr>'+availRow;
      }
      
    }else{
      availRow= availRow = '<tr><td align="left" style="border-left:solid 1px #333;border-bottom:solid 1px #333;">'+
          "None"+'</td><td align="left"  style="border-left:solid 1px #333;border-bottom:solid 1px #333;border-right:solid 1px #333;">'+
          '</td></tr>';
    }
  
    Logger.log("Avail Length ========= "+availLocArr + "============"+arr);  
    
    var partRow = '';
    if(partArr.length > 0){
      for(k=1 ;k < partArr.length ;k++){     
        for(l in namesArr){    
          if(partArr[k].toString().equals(namesArr[l].toString())){ 
            Logger.log("Name and index " + partArr[k] +"======"+(parseInt(l)+3));
            partAvailLocArr[0] = "Loc";
            partAvailLocArr[k] = sheet.getRange(parseInt(l)+3,6).getValue();  
            partAvailVal[0] ="Value";
            partAvailVal[k]=sheet.getRange(parseInt(l)+3,colIndex).getValue();
          }
        }      
      }
    
      for(y=1;y<partArr.length;y++){  
        partResData = partArr[y] +"  "+partAvailLocArr[y]+" "+ partAvailVal[y] + "; "+ partResData; 
        partRow = '<tr><td align="left" style="background:#edecec; border-left:solid 1px #747d86;border-bottom:solid 1px #747d86;">'+
          partArr[y]+'</td><td align="left" style="background:#edecec; border-left:solid 1px #747d86;border-bottom:solid 1px #747d86;">'+
          partAvailLocArr[y]+'</td><td align="left" style="background:#edecec; border-left:solid 1px #747d86;border-bottom:solid 1px #747d86;border-right:solid 1px #747d86;">'+
          partAvailVal[y]+'</td></tr>'+partRow;
      }
    }else{
      partRow = '<tr><td align="left" style="background:#edecec; border-left:solid 1px #747d86;border-bottom:solid 1px #747d86;">'+
          "None"+'</td><td align="left" style="background:#edecec; border-left:solid 1px #747d86;border-bottom:solid 1px #747d86;">'+
          '</td><td align="left" style="background:#edecec; border-left:solid 1px #747d86;border-bottom:solid 1px #747d86;border-right:solid 1px #747d86;">'+
          '</td></tr>';
    }
    Logger.log("Part Avail Length ========= "+partAvailLocArr +"============="+partAvailVal +"============"+partArr);  
  }  
  
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
'              <td valign="middle" style="background-color:#29abe0;text-align:center;font-family:\'open sans\',\'Helvetica Neue\',Arial,sans-serif;font-size:26px;text-align:center;color:#fff;padding:5px 0 !important">'+stream+' - Resource Availability</td>'+
'            </tr>'+
'            <tr>'+
'              <td><table width="620" border="0" cellspacing="0" cellpadding="0" style="margin-bottom:20px">'+
'                  <tr>'+
'                    <td width="235" align="left" valign="top"><table width="230" cellspacing="0" cellpadding="3">'+
'                      <tr>'+
'                        <td colspan="2" align="center" style="background:#2e2e2e; border: solid 1px #333; border-top:none"><span style="font-size:13.0pt;color:#fff">Available </span></td>'+
'                      </tr>'+
'                      <tr>'+
'                        <td width="153" align="left" style="background:#444f5c;  border-left:solid 1px #333;border-bottom:solid 1px #333; color:#fff"><strong>Name</strong></td>'+
'                        <td width="63" align="left" style="background:#444f5c; border-left:solid 1px #333;border-bottom:solid 1px #333; color:#fff"><strong>Location</strong></td>'+
'                        </tr>'+availRow+
'                    </table></td>'+
'                    <td width="385" align="right" valign="top"><table width="380" cellspacing="0" cellpadding="3">'+
'                      <tr>'+
'                        <td colspan="3" align="center" style="background:#646464; border: solid 1px #747d86; border-top:none"><span style="font-size:13.0pt;color:#fff">Partially Available</span></td>'+
'                      </tr>'+
'                      <tr>'+
'                        <td width="203" align="left" style="background:#747d86; border-left:solid 1px #333;border-bottom:solid 1px #333;border-bottom:solid 1px #333; color:#fff"><strong>Name</strong></td>'+
'                        <td width="74" align="left" style="background:#747d86;border-left:solid 1px #333;border-bottom:solid 1px #333; color:#fff"><strong>Location</strong></td>'+
'                        <td width="83" align="left" style="background:#747d86;border-left:solid 1px #333;border-bottom:solid 1px #333;border-right:solid 1px #333; color:#fff"><strong>Availability</strong></td>'+
'                      </tr>'+partRow+
'                    </table></td>'+
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
  
 MailApp.sendEmail(recipientsTO, subject, message, { htmlBody: html});       
        
}

function getUniqueResultantArray(uniqueTechGrp,TechGrpArr,ResArr){
  
  var arr =[];

  for(j in uniqueTechGrp){
    var tempArr =[];
    for(k in TechGrpArr){  
      if(!uniqueTechGrp[j].toString().equals("") && uniqueTechGrp[j].toString().equals(TechGrpArr[k].toString())){
        tempArr[0] = TechGrpArr[k];
        tempArr.push(ResArr[k]);
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



function sendReportInMail(e){
  
  var recipientsTO= "kalyani.kolimali@mtuity.com";
   //*****var recipientsTO = "kalyani.kolimali@mtuity.com"+","+"vijay@mtuity.com";
  var message = "Hi Vijay/Kalyani, "
  var subject = "Res Availability Failure Report";
  var html = '<body>'
       + '<p>'+message+'</p>'
       + '<p> Res Availability Failure Report Details :</p>'
       + '<p> Error Message : '+e.message+'.'+'</p>' 
       + '<p> FileName : '+e.fileName +'.'+'</p>'
       + '<p> Line Number : '+e.lineNumber +'.'+'</p>'
       +'<p> </p>'
       +'<p> </p>'
       +'<p>Thanks & Regards,<br></br>Mtuity ResAlloc Tracking Team.</p>'
       +'</body>'
       MailApp.sendEmail(recipientsTO, subject, message, { htmlBody: html});
  
}


function buildSheetName(monthVal){
  
  var currentYear = (new Date()).getYear();    
  var monthTabs = [ "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December" ];
  var sheetNameVal = monthTabs[monthVal]+" "+currentYear;  
  return sheetNameVal;
  Logger.log("SheetNameValue =========", sheetNameVal);
}

