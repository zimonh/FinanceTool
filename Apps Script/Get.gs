Logger.log('----------------------Get START');

//Settings
var col_bill           = 3;
var col_description    = 5;
var col_budget_nr      = 7;
var col_budget_price   = 8;
var col_budget_total   = 9;
var col_bill_nr        = 10;
var col_bill_price     = 11;
var col_bill_total     = 12;
var start_col_dates    = 13;
var settings_length    = 54  +1;
var col_projects_sum   = 7;
var sheet_name_settings = 'Settings';
var gsp = PropertiesService.getScriptProperties();

//Spreadsheet
var spreadsheet    = SpreadsheetApp.getActiveSpreadsheet();
function gsbn(sname){return spreadsheet.getSheetByName(sname);}

//Get_all_settings
function get_all_settings(){
var sheet_settings = gsbn(sheet_name_settings);
var all_settings = sheet_settings.getRange(1,2,settings_length,2).getValues();
gsp.setProperty('settings',array_encrypt(all_settings));}



var all_client_data = '';
var all_costs = '';

var bill = '';
var signature = '';
var doc = '';
var body = '';
var name = '';

Logger.log('------------------------Get 1');
//Build_Array_Setting START
if(gsp.getProperty('settings') !== null){
//this takes .05 sec so it could be run on opening and then only on changing settings but that is not a priority now..
  var all_settings = array_decrypt(gsp.getProperty('settings'));
  function set_a(start_row, start_col, nr_rows, nr_cols){ var arrayy = [];
    for (i = 0; i < nr_rows; i++){arrayy.push([]);
      for (j = 0; j < nr_cols; j++){arrayy[i].push(all_settings[start_row-1+i][start_col-2+j]);}}return arrayy;}
Logger.log('------------------------Get 1.1');  
  var po = 1;var ro = 0;
  ro = 4; var settings_green  = set_a(po, 2, ro, 1); po += ro;
  ro = 8; var settings_red    = set_a(po, 2, ro, 2); po += ro;
  ro = 4; var settings_purple = set_a(po, 2, ro, 1); po += ro;
  ro = 5; var settings_yellow = set_a(po, 2, ro, 1); po += ro;
  ro = 5; var settings_blue   = set_a(po, 2, ro, 1); po += ro;
  ro = 2; var settings_gray   = set_a(po, 2, ro, 2); po += ro;
  ro = 10;var settings_white  = set_a(po, 2, ro, 1); po += ro;
  ro = 1; var settings_pink   = set_a(po, 2, ro, 2); po += ro+1;
  ro = 2; var settings_file   = set_a(po, 2, ro, 1); po += ro+1;
  ro = 3; var settings_sname  = set_a(po, 2, ro, 1); po += ro+1;
  ro = 5; var settings_orange = set_a(po, 3, ro, 1); po += ro;
  ro = 1; var settings_filena = set_a(po, 2, ro, 1); po += ro;
  ro = 1; var settings_font   = set_a(po, 2, ro, 1);
  ///Build_Array_Setting END
Logger.log('------------------------Get 1.2');  
  
  
  
  
  
  
  
  
  
  function get_all_client_data(){
    Logger.log('----------------------Set all Clients START');
    
    all_client_data = gsbn(settings_sname[2][0]).getDataRange().getValues();
    set_properties('clients',all_client_data);
    //gsp.setProperty('clients',array_encrypt(all_client_data));
    Logger.log('----------------------Set all Clients END');}

Logger.log('------------------------Get 1.3');

  function get_all_costs(){
    Logger.log('----------------------Set all Costs START');
    all_bill_data = gsbn(settings_sname[0][0]).getDataRange().getValues();
    set_properties('costs',all_bill_data);
    //gsp.setProperty('costs','1')
    //for the settings properties are fine but for sheets witth lots of changes this is not.    
    //perhps in a small sheet this is fine but with over a certain number its not good
    // 9kB per value with a max of  500kB          9kb is around 8767 characters
    // there is a limit on the number of read/write operations per day: 50,000 for consumer accounts and 500,000 for G Suite Basic/Business/Edu/Gov.
   // gsbn(settings_sname[0][0]).getDataRange().getValues()
    
    /*
    //should fill costs based limit
    var all_costs_e = array_encrypt(all_costs); 
    var half_length = Math.ceil(all_costs_e.length / 2);      
    gsp.setProperty('costs',all_costs_e.slice(0,half_length))
    gsp.setProperty(('costs'+1),all_costs_e.slice(half_length));  
    */
    Logger.log('----------------------Set all Costs END');}

Logger.log('------------------------Get 1.4');   
  
  bill      = settings_file[0][0];
  signature = settings_file[1][0];


}





function onOpen(){
  if(gsp.getProperty('settings') === null){get_all_settings();
  }else{get_all_client_data();get_all_costs();URLupdate();URLupdate2();}
  
  
  
  var string = spreadsheet.getName();
  var expr = / By ZIMONH/;
  
  
    if(!expr.test(string)){string +=  ' By ZIMONH';}
    
  spreadsheet.setName(string);
  
  
  
  
}

Logger.log('------------------------Get 2');



//Build_Array_Clients
var all_client_data = '', client_length = '';
  //if(gsp.getProperty('clients') !== null){
  
  all_client_data = array_decrypt(get_properties('clients'));  
  client_length   = all_client_data.length;
  //Logger.log(all_client_data);
  //}
  
//Build_Array_Costs
var all_bill_data = '', costs_length = '';
  if(gsp.getProperty('costs1') !== null){
  
  //should be part of loop once the loop does not find any properties it should stop
  all_bill_data   = array_decrypt(get_properties('costs'));
  
  //all_bill_data = gsbn(settings_sname[0][0]).getDataRange().getValues();
  costs_length    = all_bill_data.length;
  }




var client = '';                         //this will be the filled with client info based on the active cell
var costs = [];                          //this will be filled with all the cost that match this bill number


function get_properties(varname){
var counter3 = 1;
  var stringy = '';
  var property = 1;
  while(property !== null){
    property =  gsp.getProperty(varname+counter3);
	Logger.log('Got properties: '+varname+counter3);	 
    stringy += gsp.getProperty(varname+counter3);
    counter3++;
  }
return stringy;
}


function delete_properties(varname,start){
var counter = start;
  //empty old settings
  while(gsp.getProperty(varname+counter) !== null){
  	Logger.log('Deleted properties: '+varname+counter);	 
    gsp.deleteProperty(varname+counter);
    counter++;
  }
}

function set_properties(varname,data_insert){
      var string = array_encrypt(data_insert);
      //part length
      var parts_size = 8500;
      var array_length_now = string.length;      
      var array_length_now2 = string.length; 
      var stringy = '';
      
      
      var counter4 = 1;
       while(array_length_now2-parts_size+parts_size > 0){
          counter4++;
          array_length_now2 = array_length_now2-parts_size;
      }
      
      delete_properties(varname,counter4);
      
      
      /*shoud use:
      
      .setProperties({
   'cow': 'moo',
   'sheep': 'baa',
   'chicken': 'cluck'
 });

      
      */
      
      
      //fill local settings with parts of the string
      var counter2 = 1;
      while(array_length_now-parts_size+parts_size > 0){
       var part =
       string.substring((counter2*parts_size-parts_size),
          counter2*parts_size);
      
          Logger.log('Set properties: '+varname+counter2);
          var costs_part = "";
          gsp.setProperty(varname+counter2,part);
          counter2++;
          array_length_now = array_length_now-parts_size;
      }
      
      
    
    };




function array_encrypt(array){var string = '';
for(var i = 0; i < array.length; i++){string += array_encrypt_lv(array[i])+'♣';}


return string.slice(0,-1).replace(/♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♣/g, "♠");}

function array_encrypt_lv(array){var string = '';
if(typeof array !==  "string"){
for(var i = 0; i < array.length; i++){string += array[i]+'♦';}
return string.slice(0,-1);
}else{return array;}}


function array_decrypt(string){var string = string.replace(/♠/g, "♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♣"); var array = []; var arrays = string.split('♣');
for(var i = 0; i < arrays.length; i++){
array.push(arrays[i].split('♦'));}
return array;}
Logger.log('----------------------Get END');
