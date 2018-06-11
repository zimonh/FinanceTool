//connected to a custom trigger since the default is not authorized to trigger edit of document
function onEdit(e){
  Logger.log('----------------------Edit START');

  var range = e.range;
  var active_sheet = e.source.getSheetName();

  Logger.log('----------------------Active Sheet: '+active_sheet);
  Logger.log('----------------------'+settings_sname[0][0]);
  
  if(active_sheet === settings_sname[0][0]){
    if(range.getRow() > 3){
      if(range.getColumn() >= start_col_dates && range.getColumn() < start_col_dates+4 ){
        spread(range.getRow(),range.getColumn(),e.value);
        Logger.log('----------------------Create Doc');
        
        createDoc(range.getRow(),range.getColumn(),settings_green[range.getColumn()-start_col_dates][0],e.value);
        get_all_costs();
      }else{
        get_all_costs();
      }
    }
  }
  
    if(active_sheet === settings_sname[2][0]){
    if(range.getRow() > 2){
     
        get_all_client_data();
      
    }
  }
  
  if(active_sheet === sheet_name_settings){
    if(range.getColumn() > 1){
       get_all_settings();    
       Logger.log('Settings');
    }
  }
  Protected(range, false, active_sheet);
  Logger.log('----------------------Edit END');
}
