var abc = ("ABCDEFGHIJKLMNOPQRSTUVWXYZ").split("");
 
 
 
function URLupdate(){
  var sn = gsbn(settings_sname[1][0]); //enter your personal sheet name here
  var d = new Date();
  //var n = d.getMinutes();
  var A = sn.getRange("A1");  //update the cell reference here
  var B = sn.getRange("E1");
 // var linky = "https://zimonh.at/update.html?"+n;
  //var valuu = IMPORTHTML(linky,"list",1);
  A.setValue('=IMPORTHTML("http://zh_finance_sheet_notification.accentchanger.com?'+d+'";"list";1)');  //enter your personal URL name here
  B.setValue('zh_finance_sheet_update.accentchanger.com');
}
function URLupdate2(){
  var d = new Date();
  var A = gsbn(settings_sname[0][0]).getRange("J500"); 
  A.setValue('=IMPORTHTML("http://zh_finance_sheet_notification.accentchanger.com?'+d+'";"list";2)');  //enter your personal URL name here
}






function dutchtax(totaly){
 var p1 = 2500,
     p2 = 2500,
     p3 = 5000,
     p4 = 190000,
     p5 = 200000,
     r1 = 15,
     r2 = 10,
     r3 = 5,
     r4 = 1,
     r5 = 0.5,
     payhight = 0,
     t5 = Math.max(0,  (totaly-p5) /100*r5   ),
     t4 = Math.max(0,  (Math.min(totaly-(p1+p2+p3),   p4)) /100*r4), //1900
     t3 = Math.max(0,  (Math.min(totaly-(p1+p2),      p3)) /100*r3), //250
     t2 = Math.max(0,  (Math.min(totaly-p1,           p2)) /100*r2), //250
     t1 = Math.max(0,  Math.min(totaly,               p1)  /100*r1); //375
 if(totaly<267){return 40;}
 return t1+t2+t3+t4+t5;
}


//function add(a,b){ return a + b; }

function precisionRound(number, precision) {
  var factor = Math.pow(10, precision);
  return Math.round(number * factor) / factor;
}
