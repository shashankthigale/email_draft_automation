function createBulkPDFs(){
  const docFiles = DriveApp.getFilesByName("Renewal Invoice - Template");
  
  //while (invoiceDoc.hasNext()) {
const invoiceDoc = docFiles.next();

  //console.log(invoiceDoc)
  const tempFolder = DriveApp.getFolderById("1Y5AFPu9C9BqdmCt__T5IGN7kAGt7QcVV");

  const pdfFolder = DriveApp.getFolderById("1i5YhT8eRmct4UjMEB6iiIlpUv4tDSkjG");

  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("invoiceSheet");

  //const data = currentSheet.getRange(2,1,150,23).getDisplayValues();
  const data = currentSheet.getRange(2,1,50,16).getDisplayValues();
 
  let errors = [];

  let InvoiceNo, Date, to, IUCount, PluginPrice,  total, SupportPlanPrice, pdfName, IUText;

  var PlanName = "", PlanName2 = "", PlanName3 = "";

  data.forEach(row => {
    try{
      InvoiceNo = row[12];
      Date = row[15];
      to = row[0];      
      PlanName = row[1];
      IUCount = row[5];
      PluginPrice = row[14];
      PluginRenewalAmount = row[8];
      IUText = row[13]


      SupportPlanPrice = row[10];

      total = row[11];
      console.log("total: "+total)

      pdfName = to;
      
      createPDF(InvoiceNo,Date,to,PlanName,IUCount,IUText,PluginPrice,PluginRenewalAmount,SupportPlanPrice, total, pdfName, invoiceDoc, tempFolder, pdfFolder);
      errors.push(["success"]);
    }
    catch(err){
      console.log([err]);
    }
    
  });


      //console.log(currentSheet.getRange(2,1,4,22).getDisplayValues());
createEmailDrafts();
} 

function createPDF(InvoiceNo,Date,to,PlanName,IUCount,IUText,PluginPrice,PluginRenewalAmount,SupportPlanPrice, Total, pdfName, invoiceDoc, tempFolder, pdfFolder) {
  
  const tempFile =  invoiceDoc.makeCopy(tempFolder);

  const tempinvoiceDoc = DocumentApp.openById(tempFile.getId());

  const body = tempinvoiceDoc.getBody();

  body.replaceText("{InvoiceNo}", InvoiceNo);
  body.replaceText("{date}", Date);
  body.replaceText("{to}", to);
  

  if(PlanName !== ""){
    body.replaceText("{PlanName}", PlanName);
    body.replaceText("{IUText}", IUText);
    body.replaceText("{IUCount}", IUCount);
    

    body.replaceText("{PluginPrice}",PluginPrice);
    body.replaceText("{PluginRenewalAmount}", "$"+PluginRenewalAmount);
  }
  
  body.replaceText("{SupportPlanPrice}", "\n$"+SupportPlanPrice);
  
  body.replaceText("{total}", Total);


  tempinvoiceDoc.saveAndClose();

  const pdfContentBlob = tempFile.getAs(MimeType.PDF);

  //console.log(pdfFolder)

  pdfFolder.createFile(pdfContentBlob).setName(pdfName);

  tempFolder.removeFile(tempFile);

}

function createEmailDrafts() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sendEmails"); 

  var dataRange = sheet.getDataRange();

  var data = dataRange.getValues();
  //console.log(data.length)

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    var emailAddress = row[0];
    var alternateEmailAddress = row[1]; 
    var subject = row[2]; 
    var emailBody = row[3]; 
    var body = '';

    var file = DriveApp.getFilesByName(emailAddress);
    var fileName = file.next()
    var blob = fileName.getBlob();
    blob.setName("Renewal Invoice.pdf")
    
    if (emailAddress) {

      if (alternateEmailAddress) {
        GmailApp.createDraft(emailAddress, subject, body, {cc:'shashank@xecurify.com;' + alternateEmailAddress, htmlBody:emailBody, attachments:[blob]});
      }
      else
      {
        GmailApp.createDraft(emailAddress, subject, body, {cc:'shashank@xecurify.com', htmlBody:emailBody, attachments:[blob]}); 
      }

      console.log("Draft created for: " + emailAddress + " with CC: " + alternateEmailAddress);
    }
  }
}
