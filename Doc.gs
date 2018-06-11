function createDoc(row,col,bill_type,val){

  Logger.log('----------------------Create doc START');
  Logger.log('Document type: '+bill_type);


//Needs to be here
  var  doc = DocumentApp.openById(settings_file[0][0]);
  var  body = doc.getBody();
  var A = gsbn(settings_sname[1][0]).getRange("A1").getValue();  //update the cell reference here
if(A == "You have the latest version"){
  var  name = doc.getName();


 var body_index = 0,
     active_row = all_bill_data[row-1],
     time = new Date(),
     the_time = ' ' + ("0" + time.getHours()).slice(-2) + ":" + ("0" + time.getMinutes()).slice(-2),
     costs_table= [],
     total_price = [];


var date = new Date(Date.UTC(1899, 11, 30, 0, 0, val * 86400));
var datte = Utilities.formatDate(new Date(date.toUTCString()), "GMT+0700", "yyyy-MM-dd").replace("1970-01-01","")


Logger.log('Connected to Document');


 //get curent client
 for (var i=1; i <=client_length-1; i++){
  if (all_client_data[i][0] == active_row[1]){
  client = all_client_data[i];
  }
 }
 


 body.clear();
Logger.log('Empty the document');


 //Build contact info array
 var contact_table = [[
     client[0]+'\n'+     //Client Company
     client[1]+'\n'+     //Client Name
     client[2]+'\n'+     //Client Street
     client[3]+'\n'+     //Client Postcode
     client[4]+'\n'+     //Client City
     client[5]           //Client Country
     ,
     settings_red[0][0]+':\n'+  //Company:
     settings_red[1][0]+':\n'+  //Website:
     settings_red[2][0]+':\n'+  //Contact:
     settings_red[3][0]+':\n'+  //Street:
     settings_red[4][0]+':\n'+  //Postcode:
     settings_red[5][0]+':\n'+  //Email:
     settings_red[6][0]+':\n'+  //Tax nr:
     settings_red[7][0]+':'     //Company nr:
     ,
     settings_red[0][1]+'\n'+   //Company name
     settings_red[1][1]+'\n'+   //Website
     settings_red[2][1]+'\n'+   //Contact name
     settings_red[3][1]+'\n'+   //Street
     settings_red[4][1]+'\n'+   //Postcode
     settings_red[5][1]+'\n'+   //Email
     settings_red[6][1]+'\n'+   //Tax nr
     settings_red[7][1]         //Company nr
    ]];


 //insert contact info
 body.insertTable(body_index,contact_table)
     .setBorderColor('#ffffff')
     .setColumnWidth(0, 300)
     .setColumnWidth(1, 59);body_index++;

Logger.log('_______Added - Contact table');


  //Add title with Bill type
  body.insertParagraph(body_index,bill_type)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);body_index++;

Logger.log('_______Added - Title');


   var bill_dates = '';
  //If this is a regular Bill
  if(bill_type===settings_green[0][0]){
   bill_dates =
   settings_purple[1][0]+': '+                                            // Date:
   datte+'\n';                      //-BillDate-
  }

  //If this is a REMINDER
  if(bill_type===settings_green[1][0]){
   //Bill dates
   bill_dates =
   settings_purple[1][0]+': '+                                          // Date:
   datte+'\n'+                     //-Date-
   settings_purple[2][0]+': '+                                          // Bill Date:
   Utilities.formatDate(new Date(active_row[col-2]), "GMT+0700", "yyyy-MM-dd").replace("1970-01-01","")+'\n';                    //-Bill Date-
 }
 var sign = false;
 //If this is a Final Reminder
 if(bill_type===settings_green[2][0]){
   bill_dates =
   settings_purple[1][0]+': '+                                            // Date:
   datte+'\n'+                      //-Date-
   settings_purple[3][0]+': '+                                            // Reminder Date:
   Utilities.formatDate(new Date(active_row[col-2]), "GMT+0700", "yyyy-MM-dd").replace("1970-01-01","")+'\n'+                      //-Reminder Date-
   settings_purple[2][0]+': '+                                            // Bill Date:
   Utilities.formatDate(new Date(active_row[col-3]), "GMT+0700", "yyyy-MM-dd").replace("1970-01-01","")+'\n';                      //-Bill Date-
      var signature_get = DriveApp.getFileById(signature).getBlob();
   sign = true;
 }

 //If this is a Overdue bill
 if(bill_type===settings_green[3][0]){
   bill_dates =
   settings_purple[1][0]+': '+                                            // Date:
   datte+'\n'+                      //-Date-
   settings_green[2][0]+': '+                                             // Final Reminder:
   Utilities.formatDate(new Date(active_row[col-2]), "GMT+0700", "yyyy-MM-dd").replace("1970-01-01","")+'\n'+                      //-Final Reminder-
   settings_purple[3][0]+': '+                                            // Reminder Date:
   Utilities.formatDate(new Date(active_row[col-3]), "GMT+0700", "yyyy-MM-dd").replace("1970-01-01","")+'\n'+                      //-Reminder Date-
   settings_purple[2][0]+': '+                                            // Bill Date:
   Utilities.formatDate(new Date(active_row[col-4]), "GMT+0700", "yyyy-MM-dd".replace("1970-01-01",""))+'\n';                      //-Bill Date-
   var signature_get = DriveApp.getFileById(signature).getBlob();
   sign = true;
 }

//Bill Number
bill_dates += settings_purple[0][0] + ': ' + active_row[2]+'\n';





//Add general info about Bill dates
body.insertParagraph(body_index,bill_dates);body_index++;
Logger.log('_______Added - Dates');


//Logger.log(costs_length-4+'__'+all_bill_data.length);




//Should run before
//in the cost sheet find the rows that maches the filter
var costs = [];
//var all_bill_data = get_costs(4,col_description,costs_length-4,7);






//var costs_rows = get_costs(4,col_description,sheet_costs.getNumRows(),3);

for(var i=0; i <= costs_length-5; i++){
  if(all_bill_data[i][2] == active_row[2]){
  costs.push(all_bill_data[i]);
  }
}

Logger.log('Costs Array build');



  //Add header of costs
  costs_table.push([
    settings_yellow[0][0],  //Description
    settings_yellow[1][0],  //Type
    settings_yellow[2][0],  //Amount
    settings_yellow[3][0],  //Price per piece
    settings_yellow[4][0]]);//Total
  for(var j=0; j < costs.length; j++){
    total_price.push(costs[j][col_bill_total-1]); //add the prices to the total price array
    costs_table.push([
    costs[j][col_description-1],     //-Description-
    costs[j][col_description],       //-Type-
    costs[j][col_bill_nr-1],         //-Amount-
    settings_blue[0][0]+             //Euro symbol
    costs[j][col_bill_price-1],      //-Price per piece-
    settings_blue[0][0]+             //Euro symbol
    costs[j][col_bill_total-1]]);    //-Total-
   }
   
   
   
   
   var total_price_sum=0;
   var total_price_min=0;
   
   
   
   var i=total_price.length;
   while(i--){
     var total = parseFloat(total_price[i]);
     if(total<0){
       total_price_min += total;
     }else{
       total_price_sum += total;
     }
   }

  body.insertTable(body_index,costs_table)
      .setBorderColor('#ffffff')
      .setColumnWidth(0, 180)
      .setColumnWidth(1,65)
      .setColumnWidth(2,65)
      .setColumnWidth(2,65);body_index++;
  Logger.log('_______Added - Costs Table');


//Adjust pagebreak_space based on type
var pagebreak_space = 3;    //Max nr of rows before page break
if(bill_type===settings_green[1][0]){
 pagebreak_space = 0;
}

 var fee_text1 = '';
 var fee_text2 = '';
 var fee = 0;
 var late = 0;

 
if(bill_type===settings_green[2][0]){
 pagebreak_space = 0;
 fee =  parseFloat(Math.round(dutchtax(total_price_sum + total_price_min) * 100) / 100).toFixed(2);
 
}

if(bill_type===settings_green[3][0]){
 pagebreak_space = 1;
   fee =  parseFloat(Math.round(dutchtax(total_price_sum + total_price_min) * 100) / 100).toFixed(2);
   fee_text1 = '\n'+settings_blue[4][0];      //Fee
   fee_text2 = '\n'+settings_blue[0][0]+fee;  //-Fee-
   late = fee;
}

 

 //if its a big bill add a pagebreak
  if(body.getTables()[1].getNumRows()>pagebreak_space+1 && body.getTables()[1].getNumRows()!==16){
   body.insertParagraph(body_index,'').appendPageBreak();body_index++;
   body.insertParagraph(body_index,'\n\n\n\n\n');body_index++;
     Logger.log('_______Added - Page break');
  }





 var total_price_table = [[
  settings_yellow[4][0]                                    //Total header
  ,
  settings_blue[3][0]+' '+ settings_blue[1][0]+'\n'+       //Excluding Tax
  '('+settings_blue[2][0]+'%) '+settings_blue[1][0]+       //(21%) Tax
  fee_text1                                                //Fee
  ,
  settings_blue[0][0]+                                                                                        //Euro symbol
  parseFloat(Math.round(total_price_sum * 100) / 100  + total_price_min).toFixed(2)+'\n'+                     //-Excluding Tax amount-
  settings_blue[0][0]+                                                                                        //Euro symbol
  parseFloat(Math.round((total_price_sum + total_price_min) * (settings_blue[2][0]/100) * 100) / 100).toFixed(2)   +              //-Tax amount-
  fee_text2                                                                                                   //-Fee amount-
 ]];

 //Insert Subtotals
 body.insertTable(body_index,total_price_table)
     .setBorderColor('#ffffff')
     .setColumnWidth(0,257)
     .setColumnWidth(1,100)
     .setColumnWidth(2,100);body_index++;

  Logger.log('_______Added - Total Table');


 //Insert line
 body.insertHorizontalRule(body_index);body_index++;


  Logger.log('_______Added - Horizontal ruler');
 //Logger.log((total_price_sum + total_price_min)*(1+(settings_blue[2][0]/100))+Number(late));
 //Insert Total price
 body.insertParagraph(body_index,
       settings_blue[0][0]+//Euro symbol
       precisionRound( (total_price_sum + total_price_min)*(1+(settings_blue[2][0]/100))+Number(late)  ,2)) //-total_price including tax-
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
      .setSpacingBefore(0);body_index++;

  Logger.log('_______Added - Total Paragraph');



  var letter_end_text1 = '',
      letter_end_text2 = '',
      letter_end_text3 = '';

  //If this is a regular Bill
  if(bill_type===settings_green[0][0]){
    letter_end_text1 = settings_orange[0][0];                                  //Please trancfer within .. days,
  }

  //If this is a REMINDER
  if(bill_type === settings_green[1][0]){
    letter_end_text1 = settings_orange[1][0]+'\n\n'+settings_orange[2][0];     //According to our administration you have not payed please pay within 3 days
    letter_end_text2 = '\n';
    letter_end_text3 = settings_white[9][0]+'\n';                              //If you already payed please gnore this letter
 }
 
 //If this is a Final notice
 if(bill_type === settings_green[2][0]){
    letter_end_text1 = settings_white[3][0]+'\n\n'+settings_orange[3][0];       //I have already asked you several times to pay the invoice above, so far I have not received any payment yet.
    letter_end_text2 = '\n\n'+
    settings_white[7][0]+' '+                                                   //If this bill is not payed in time there will be a
    settings_blue[0][0]+                                                        //Euro Symbol
    fee+' '+                                                                    //-Max Fee if not payed-
    settings_white[8][0]+'\n';                                                  //Billing fee will be charged
    letter_end_text3 = settings_white[9][0]+'\n';                               //If you already payed please gnore this letter
 }
 //Incasso
 if(bill_type === settings_green[3][0]){
    letter_end_text1 = settings_orange[4][0];                              //Please trancfer within .. days,
    letter_end_text3 = '\n'+settings_white[9][0]+'\n';                               //If you already payed please ignore this letter
 }

//Insert End text
body.insertParagraph(body_index,
  '\n\n'+
  letter_end_text1+'\n\n'+
  settings_white[5][0]+':\n'+ //using payment info:
  settings_white[6][0]+': '+  //Payment idnumber:
  '0000000000000000'.slice(String(active_row[2]).length)+active_row[2]+'\n'+
  settings_gray[0][0]+': '+   //Acount number:
  settings_gray[0][1]+'\n'+   //-Account number-
  settings_gray[1][0]+': '+   //Beneficiery:
  settings_gray[1][1]+        //-Beneficiery-
  letter_end_text2+'\n'+
  letter_end_text3+'\n'+      //If you already payed pleae ignore this letter
  settings_pink[0][0]+'\n'+   //Friendly greetings
  settings_pink[0][1]         //-From-
  );body_index++;


Logger.log('_______Added - Final paragraph');

//Insert Signature
if(sign){
  function getimg(maxwidth,position,body_index){
     body.insertParagraph(body_index,'&Logo&');body_index++;
     
     Logger.log('_______Added - IMG Paragraph');
     
     var image = DriveApp.getFileById(settings_file[1][0]).getBlob();
         oImg = body.findText("&Logo&").getElement().getParent().asParagraph();
     Logger.log('IMG ready');
     
     if(position === 'r'){oImg.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);}
     if(position === 'c'){oImg.setAlignment(DocumentApp.HorizontalAlignment.CENTER);}
     if(position === 'l'){oImg.setAlignment(DocumentApp.HorizontalAlignment.LEFT);}
    
     oImg.clear();
     Logger.log('Empty the IMG Paragraph');
     oImg = oImg.appendInlineImage(image);
     Logger.log('_______Added - IMG');
     var ww = oImg.getWidth(),
         hh = oImg.getHeight(),
         of = ww / maxwidth;
     
     oImg.setWidth(ww / of);
     oImg.setHeight(hh / of);
     Logger.log('IMG Width set');
  }
  getimg(100,'l',body_index);
}







 //STYLING-------------------------------------//

//get all tables
var tables = body.getTables();

//contact_table
tables[0]
 .getCell(0,1)
 .editAsText()
 .setFontSize(9);

tables[0]
 .getCell(0,0).getChild(0)
 .editAsText()
 .setFontSize(16).setBold(true);
  

tables[0]
 .getCell(0,2)
 .editAsText()
 .setFontSize(9);



//costs_table
tables[1]
 .getRow(0)
 .editAsText()
 .setBold(true);

for(var q = 0; q<5; q++) {
 tables[1]
  .getCell(0, q)
  .setBackgroundColor('#000000')
  .editAsText()
  .setForegroundColor('#FFFFFF');
}

for(var q = 0; q < tables[1].getNumRows(); q++) {
 tables[1]
  .getCell(q,1)
  .getChild(0)
  .asParagraph()
  .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
 tables[1]
  .getCell(q,2)
  .getChild(0)
  .asParagraph()
  .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

 tables[1]
  .getCell(q,3)
  .getChild(0)
  .asParagraph()
  .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
 tables[1]
  .getCell(q,4)
  .getChild(0)
  .asParagraph()
  .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
}

//total_price_table
tables[2]
 .getCell(0,0)
 .getChild(0)
 .asParagraph()
 .setHeading(DocumentApp.ParagraphHeading.HEADING1)
 .setSpacingBefore(0)
 .setSpacingAfter(0);

tables[2]
 .getCell(0,1)
 .getChild(0)
 .asParagraph()
 .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

tables[2]
 .getCell(0,1)
 .getChild(1)
 .asParagraph()
 .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

if(bill_type===settings_green[3][0]){
tables[2]
 .getCell(0,1)
 .getChild(2)
 .asParagraph()
 .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

tables[2]
 .getCell(0,2)
 .getChild(2)
 .asParagraph()
 .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
}

tables[2]
 .getCell(0,2)
 .getChild(0)
 .asParagraph()
 .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

tables[2]
 .getCell(0,2)
 .getChild(1)
 .asParagraph()
 .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);




 body.setFontFamily(settings_font[0][0]);

 filename = settings_filena[0][0]
 .replace("<billnr>", active_row[2])
 .replace("<billtype>", bill_type)
 .replace("<client>", client[0])
 .replace("<datetime>", datte+the_time);

 doc.setName(filename);

Logger.log('Name Set');

}else{


 body.clear();
 body.insertParagraph(1,'This version is out of date.\nYou can find the latest version: ');
 body.insertParagraph(2,'here').setLinkUrl("zh_finance_sheet_update.accentchanger.com");
}
}
