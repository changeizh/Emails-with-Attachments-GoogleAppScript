function emailer(){
   
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt('Start emailing?', ui.ButtonSet.YES_NO);
  
  
  if(response.getSelectedButton()== ui.Button.YES){
    
    
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Emails');
  
 
  cnt  =2;
  while(true){
  
  data = sheet.getRange(cnt,1).isBlank();
    
    flag = sheet.getRange(cnt, 5).getValue();
    
    if(data == false){
      
      sheet.getRange(cnt, 4).insertCheckboxes()
      
      var sub =  sheet.getRange(cnt, 1).getValue();
      
      var cmp = sheet.getRange(cnt, 2).getValue();
      
      var file_id = getFileID(sub, cnt);
      
      if(flag != 'EMAIL_SENT'){
      
        var reci = sheet.getRange(cnt, 3).getValue();
        
        var subj = sub+" - "+cmp;
        
        try{
          
          var attach =  DriveApp.getFileById(file_id);
           
          var sig = HtmlService.createHtmlOutputFromFile('signature').getContent();
         
          GmailApp.sendEmail(reci, subj,sig, {htmlBody : sig,name :"Seair Exim Data Dispatch", attachments :[attach.getAs(MimeType.MICROSOFT_EXCEL)] });
          
          sheet.getRange(cnt, 5).setBackground('green');
          
          sheet.getRange(cnt, 5).setValue('EMAIL_SENT');
          
          sheet.getRange(cnt, 4).check();
          
        }
        
        catch (e){
          
          console.log()
          
          console.error('API got an error: '+ e)
          
          var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Emails');
          
          if(e == 'Exception: Invalid argument: id'){
            
            sheet1.getRange(cnt, 4).uncheck();
            
            sheet1.getRange(cnt, 5).setValue('File does not Exist?: ' +e);
          
            sheet1.getRange(cnt, 5).setBackground('pink');
       
          }
          
          else{
            
            sheet1.getRange(cnt, 4).uncheck();
            
            sheet1.getRange(cnt, 5).setValue('API got an error: ' +e);
          
            sheet1.getRange(cnt, 5).setBackground('pink');
          
          }
          
          
        }
        
      }
      
      else{
        sheet.getRange(cnt, 4).check();
        cnt++;
        continue;
        
      }
     
       
    }
    else{
      break;
    }
   cnt++; 
  }
}
};



//------------------------------------------------------------------------------------------


function getFileID(filename, num){
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Emails');
  
  var files = DriveApp.getFiles();
  
  while (files.hasNext()) {
      
    var file = files.next();
    
    try{
      
          if(filename == file.getName()){
        
           sheet.getRange(num, 4).check();
         
           return file.getId();
         
           break;
    
           }
      
          else{
      
              continue;
    
             }
          
       }
    
    catch (err){
      
      console.log()
          
      console.error('API got an error: '+ err)
      
      var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Emails');
      
      sheet2.getRange(num, 4).uncheck();
      
      sheet2.getRange(cnt, 5).setValue('File does not Exist?: ' +err);
      
    
    
    }
    
      
  }
  
};




//------------------------------------------------------------------------------------------------------




function listFilesInFolder() {
  

  
    //if Emailer Sheet doesn;t exist, create Sheet and append 
  var em = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Emails');
  
  if (!em) {

  var mod = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Emails');
   
  mod.appendRow(["File_Name", "Company Name", "Email", "Files Exist?", "Email Status"]);
  mod.getRange(1 ,1, 1, 5).setBackground('#87CEEB');
  mod.setFrozenRows(1);
   
 }
  
  var Gdrive = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('List of Files in GDrive');
  
//if GDrive Sheet doesn;t exist, create Sheet and append 
 if (!Gdrive) {

   var Gdrive = SpreadsheetApp.getActiveSpreadsheet().insertSheet('List of Files in GDrive');
   
   Gdrive.appendRow(["Name", "Date", "Size", "URL", "Download", "Description"]);
   Gdrive.getRange(1 ,1, 1, 6).setBackground('orange');
   Gdrive.setFrozenRows(1);
   
}
   

  
  //--------------------------------
  
 
  
  var contents = DriveApp.getFilesByType(MimeType.MICROSOFT_EXCEL);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('List of Files in GDrive');
  sheet.clear();

  
  sheet.appendRow(["Name", "Date", "Size", "URL", "Download", "Description"]);

  while (contents.hasNext()) {
      
    var file = contents.next();
    
    var data = [
      file.getName(),
      file.getDateCreated(),
      (file.getSize()/1024/1024).toFixed(2) + " MB",
   
      file.getUrl(),
      "https://docs.google.com/uc?export=download&id=" + file.getId(),
      file.getDescription(),
      file.getDescription()
    ];

    sheet.appendRow(data);
         
       
  }
  
 setValidation();

  
};





//----------------------------------------------------------------------------------------------------------------------------


function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Emailer')
      .addItem('Refresh Files', 'user_interface')
      .addItem('Start Emailer', 'emailer')
      .addToUi();
}




//---------------------------------------------------------------------------------------------------

function setValidation() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Emails');
  var Gdrive = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('List of Files in GDrive');
  var cells = sheet.getRange("A2:A" + sheet.getLastRow()+100);
  var rules = Gdrive.getRange("A2:A" + Gdrive.getLastRow());
  var validation = SpreadsheetApp.newDataValidation().requireValueInRange(rules).build();
  cells.setDataValidation(validation);
  sheet.autoResizeColumns(1, 5);
  Gdrive.autoResizeColumns(1, 3);
}



//-------------------------------

// user interface desing

function user_interface(){
  var ui = SpreadsheetApp.getUi();
  
//  ui.alert('Refreshing Google Drive Items in Files Sheets');
  
  var response = ui.prompt('Refresh Google Drive Files', ui.ButtonSet.YES_NO);
  
  
  if(response.getSelectedButton()== ui.Button.YES){
    
   listFilesInFolder();
    
  }

}





