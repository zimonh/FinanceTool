function spread(row,col,val){
Logger.log('----------------------Spread START');
var dippy = [];

if(val === undefined || typeof val !== 'string'){
 val = '';
}

for(var i=0; i <= costs_length-5 ; i++){
  if(all_bill_data[i+3][2] === all_bill_data[row-1][2] && all_bill_data[i+3][2] !== ''){
    dippy.push([val]);
  }else{
    var rowval = all_bill_data[i+3][col-1];
   
    if(rowval.lenght < 15){
    dippy.push([rowval]);
    }else{
      
      if(rowval !== ''){
      rowval = Utilities.formatDate(new Date(rowval), "GMT+0700", "dd/MM/yyyy");
      dippy.push([rowval]);
      }else{dippy.push([rowval]);}
    }
  }
}


var langger = dippy.length;

gsbn(settings_sname[0][0]).getRange(4,col ,langger,1).setValues(dippy).setNumberFormat("mm-dd");

Logger.log('----------------------Spread END');
}
