

function Protected(range, m, active_sheet) {
 Logger.log('----------------------Protect START');
 var mrao = (249+251);
  
 if(active_sheet === sheet_name_settings || m){

  if(range.getColumn() === 3 || m){
    gsbn(sheet_name_settings)
    .getRange(48,3,5,1)
    .setValues([
    ['=SUBSTITUTE(B29;"<deadline>";B48)'],
    ['=SUBSTITUTE(B30;"<deadline>";B48)'],
    ['=SUBSTITUTE(B31;"<deadline>";B50)'],
    ['=SUBSTITUTE(B29;"<deadline>";B51+1)'],
    ['=SUBSTITUTE(B33;"<deadline>";B52)']
    ]);
  Logger.log('Protected 1');
  }


   if(range.getColumn() === 1 || m){
      gsbn(sheet_name_settings)
      .getRange(1,1,settings_length-1,1)
      .setValues([
     ['Bill'],
     ['Payment Reminder'],
     ['Notice of Non payment'],
     ['Bill with Fee'],
     ['Company'],
     ['Website'],
     ['Contact'],
     ['Street'],
     ['Postcode'],
     ['Email'],
     ['Tax number'],
     ['Busness number'],
     ['Billnumber'],
     ['Date'],
     ['Bill date'],
     ['Reminder date'],
     ['Description'],
     ['Type'],
     ['Amount'],
     ['Price per piece'],
     ['Total'],
     ['Curency'],
     ['Tax'],
     ['Tax rate'],
     ['Excluding'],
     ['Late Fee'],
     ['Acountnumber'],
     ['In name of'],
     ['Please transfer the total amount within <deadline> days,'],
     ['According to our records your payment is overdue. The due date was <deadline> days after the invoice date.'],
     ['Therefore I request, you pay the invoice amount above within <deadline> days after this reminder,'],
     ['I have notified you several times to pay this invoice, to my disappointment I have not received any payment yet.'],
     ['After the final notice, I have not received any payment. Therefore you will be charged a late fee. If the payment with late fee is not received legal action will be taken. Payment can be made up to <deadline> days of the date of this letter,'],
     ['Citing'],
     ['Payment reference'],
     ['If this invoice is not paid in time after this final notice a'],
     ['collection fee will be charged.'],
     ['If you have already sent your payment, please disregard this letter.'],
     ['Yours truly,'],
     ['Files'],
     ['Google Drive file id for invoice Document:'],
     ['File id for Siganture Image:'],
     ['Sheetnames   '],
     ['Cost sheet name:'],
     ['Projects sheet name:'],
     ['Clients sheet name:'],
     ['Deadlines'],
     [''],
     ['Bill'],
     ['Reminder part2'],
     ['Final notice'],
     ['Bill with late fee'],
     ['File Name\n<billnr>\n<billtype>\n<client>\n<datetime>'],
     ['Font']
      ]);
   Logger.log('Protected 2');
   }
 }



 if(active_sheet === settings_sname[1][0] || m){
  if(range.getRow() > 2 || m){
    if(range.getColumn() > 6 || m){
       
       Logger.log('Protected 3');
       var sn = settings_sname[0][0];
       
       gsbn(settings_sname[1][0])
       .getRange(3, 7, gsbn(settings_sname[1][0])
       .getMaxRows()-3,1)
       .setValue('=IF(H$3:H="";"";IF(F3:F>H3:H;F$3:F-J$3:J;H$3:H-J$3:J))');
       
       Logger.log('Protected 3.1');
       gsbn(settings_sname[1][0])
       .getRange(3, 8, gsbn(settings_sname[1][0])
       .getMaxRows()-3,1)
       .setValue('=IF(E$3:E*F$3:F=0;""; IF(SUMIF('+sn+'!A$4:A;A$3:A;'+sn+'!I4:I)=0;E$3:E*F$3:F;if(A$3:A="";"";SUMIF('+sn+'!A4:A;A$3:A;'+sn+'!I4:I))))');
       
       Logger.log('Protected 3.2');
       gsbn(settings_sname[1][0])
       .getRange(3, 9, gsbn(settings_sname[1][0])
       .getMaxRows()-3,1)
       .setValue('=IF(A$3:A="";"";COUNTIFS('+sn+'!A$4:A;A$3:A;'+sn+'!M$4:M;""))');
       
       Logger.log('Protected 3.3');
       gsbn(settings_sname[1][0])
       .getRange(3, 10, gsbn(settings_sname[1][0])
       .getMaxRows()-3,1)
       .setValue('=IF(A$3:A="";"";SUMIFS('+sn+'!L$4:L;'+sn+'!A$4:A;A$3:A;'+sn+'!M$4:M;">20"))');

    }
  }
  if(range.getRow() === 1){
    var text = [];
    for (var i =0; i < 4; i++) {
      text.push('=SUM('+abc[col_projects_sum+i-1]+'3:'+abc[col_projects_sum+i-1]+')');
      Logger.log('Protected 4_'+i);
    }
    if(range.getColumn() > 6){
        gsbn(settings_sname[1][0])
        .getRange(1, col_projects_sum, 1, 4)
        .setValues([text]);
        Logger.log('Protected 5');
    }
  }
 }
 
 
 if(active_sheet === settings_sname[0][0] || m){
   if(range.getRow() === (249+251) || m){
    
     gsbn(settings_sname[0][0])
     .getRange(500,4)
     .setValue('This version is limited to 500 rows\nfor help contact: ');
  
     
     
     gsbn(settings_sname[0][0])
     .getRange(500,1,1,17)
     .setFontSize(17)
     .setBackground("red")
     .setBackground("red")
     .setWrap(true)
     .setFontColor("white");
    }
    
    if(range.getRow() < 3 || m){
      if(range.getColumn() === 4 || m){
         gsbn(settings_sname[0][0])
         .getRange(1, 4)
         .setValue('=Lists!C2');
         
         gsbn(settings_sname[0][0])
         .getRange(2, 4)
         .setValue('=Lists!C3');
      }
    }

   if(range.getColumn() === col_budget_total || m){
    var nr = abc[col_budget_nr], pr = abc[col_budget_price], ex = abc[3];
      if(range.getRow() > 2 || m){
        gsbn(settings_sname[0][0])
        .getRange(4, col_budget_total, gsbn(settings_sname[0][0])
        .getMaxRows()-3,1)
        .setValue('=IF(H:H="";"";IF(G:G="";IF(D:D=D$1;H:H;H:H*-1);IF(D:D=D$1;H:H*G:G;H:H*G:G*-1)))');
        Logger.log('Protected 6');
      }
    }

    if(range.getColumn() === col_bill_total || m){
     var nr = abc[col_bill_nr], pr = abc[col_bill_price], ex = abc[3];
      if(range.getRow() > 2){
         gsbn(settings_sname[0][0])
         .getRange(4, col_bill_total,gsbn(settings_sname[0][0])
         .getMaxRows()-3,1)
         .setValue('=IF(K:K="";"";IF(J:J="";IF(D:D=D$1;K:K;K:K*-1);IF(D:D=D$1;K:K*J:J;K:K*J:J*-1)))');
         Logger.log('Protected 7');
      }
    }

   if(range.getRow() < 2 || m){
     if(range.getColumn() >= start_col_dates || m){
       
       gsbn(settings_sname[0][0])
       .getRange(1,start_col_dates)
       .setValue('=Settings!B'+(settings_length-7));
       
       gsbn(settings_sname[0][0])
       .getRange(1,start_col_dates+1)
       .setValue('=Settings!B'+(settings_length-5));
       
       gsbn(settings_sname[0][0])
       .getRange(1,start_col_dates+2)
       .setValue('=Settings!B'+(settings_length-4));
       
       gsbn(settings_sname[0][0])
       .getRange(1,start_col_dates+3)
       .setValue('=Settings!B'+(settings_length-3));
       Logger.log('Protected 8');
     }
   }
   
    if(range.getRow() > mrao || m){
       gsbn(settings_sname[0][0])
       .deleteRows((249+252),(gsbn(settings_sname[0][0]).getMaxRows()-500));
    }
  }
 Logger.log('----------------------Protect END');
}

