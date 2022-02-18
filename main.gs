function editFunc(e) {
  var edit_range = e.range;
  var edit_row = edit_range.getRow();
  var edit_col = edit_range.getColumn();

  // don't mess with the header
  if (edit_row < 3) {
    return false
  };

  // edit in col 1: adding new doc or 
  // edit in col 2: manual password set
  if (edit_col == 1) {
    addDoc(e, edit_row);
  } else if (edit_col == 2) {
    setCode(e, edit_row);
  };
  
};

function addDoc(e, edit_row) {
  // for ADDING a new document or  in column 1
  var sheet = e.source;
  var item_ID = e.value;
  
  // generate a code and write to sheet
  var new_code = Math.random().toString(20).substr(2,8);
  sheet.getRange("B"+edit_row).setValue(new_code);

  // generate a new form, set up its submit trigger and spreadsheet output, and write to sheet
  try {
    var item = DriveApp.getFileById(item_ID);
  }
  catch(err) {
    var item = DriveApp.getFolderById(item_ID);
  };
  var title = item.getName();
  var form = createForm(title,new_code,item_ID);
  var output_ss = SpreadsheetApp.create(title);
  var output_ss_ID = output_ss.getId();
  form.setDestination(FormApp.DestinationType.SPREADSHEET,output_ss_ID);
  var output_file = DriveApp.getFileById(output_ss_ID);
  output_file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
  
  // set up validation trigger
  ScriptApp.newTrigger('validateAndShare')
    .forForm(form)
    .onFormSubmit()
    .create();
  var form_url = form.getPublishedUrl();
  sheet.getRange("C"+edit_row).setValue(form_url);
  sheet.getRange("D"+edit_row).setValue(form.getId());
  sheet.getRange("E"+edit_row).setValue(output_ss.getUrl());
};

function createForm(title,code,item_ID) {
  var form = FormApp.create(title)
    .setTitle(title)
    .setDescription(item_ID)
    .setCollectEmail(true);
  // this is a sorta hacky way to do it that just doesn't let you submit the form if the password is incorrect but afaik you can't set a correct answer for short responses at the moment
  var code_val = FormApp.createTextValidation()
    .setHelpText('Incorrect code')
    .requireTextMatchesPattern(code)
    .build();
  var code_item = form.addTextItem()
    .setTitle('Please enter the password')
    .setRequired(true)
    .setValidation(code_val);
  return form;
};

function validateAndShare(e) {
  var item_ID = e.source.getDescription();
  //SpreadsheetApp.getActive().getRange('A20').setValue(item_ID);
  try {
    var item = DriveApp.getFileById(item_ID);
  }
  catch(err) {
    var item = DriveApp.getFolderById(item_ID);
  };
  item.addViewer(e.response.getRespondentEmail());
  // we don't need to check the response because the form doesn't let you submit unless it's right, sorta hacky but this way we avoid needing parameters on this
  // oh but we need to get the item
  // maybe we can maintain a static dictionary matching forms to s?
  // jk no that's too much effort we'll just dump the id in the form description
};

function setCode(e, edit_row) {
  var sheet = e.source;
  var new_code = e.value;
  var form_ID = sheet.getRange("D"+edit_row).getValue();
  changeCode(new_code,form_ID);
};

function changeCode(new_code,form_ID) {
  // given a new code and a form ID, change the form's code
  var form = FormApp.openById(form_ID);
  var new_code_val = FormApp.createTextValidation()
    .setHelpText('Incorrect code')
    .requireTextMatchesPattern(new_code)
    .build();
  var code_item = form.getItems()[0].asTextItem();
  code_item.setValidation(new_code_val);
};

function testChangeCode() {
  changeCode('12345','1QQ0UaDO70XaynzDXh-UuM2buLrCNRgSl3F1nvUPm9yw');
};

function timedFullRefresh(e) {
  // for proper trigger signature
  fullRefresh();
};

function fullRefresh() {
  // removes ALL viewers and changes ALL codes in spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  for (var i = 3; i < sheet.getLastRow()+1; i++){
    try {
      var item_ID = sheet.getRange("A"+i).getValue();
      var form_ID = sheet.getRange("D"+i).getValue();
      removeViewers(item_ID);
      generateNewCode(i,form_ID, sheet);
    }
    catch (err) {
      console.log('Error on row ' + i);
      console.log(err)
      continue;
    };
  };
};

function timedRefreshAccess(e) {
  // for proper trigger signature
  refreshAccess();
};

function refreshAccess() {
  // removes ALL viewers in spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  for (var i = 3; i < sheet.getLastRow()+1; i++){
    try {
      var item_ID = sheet.getRange("A"+i).getValue();
      removeViewers(item_ID);
    }
    catch (err) {
      console.log('Error on row ' + i);
      continue;
    };
  };
};

function removeViewers(item_ID) {
  try {
    var item = DriveApp.getFileById(item_ID);
  }
  catch(err) {
    var item = DriveApp.getFolderById(item_ID);
  };
  var viewers = item.getViewers();
  // console.log(typeof viewers[0])
  for (var i = 0; i < viewers.length; i++) {
    // console.log(viewers[i].getEmail());
    item.revokePermissions(viewers[i]);
  };
};

function testRemoveViewers() {
  var item_ID = '13jkvHp3IZjovHzj5h54-H9lf3he-EtwAPyMLyEp0Eo4';
  removeViewers(item_ID);
};

function generateNewCode(edit_row, form_ID, sheet) {
  // regenerate code and write new code to sheet
  var new_code = Math.random().toString(20).substr(2,8);
  sheet.getRange("B"+edit_row).setValue(new_code);
  // update form
  changeCode(new_code, form_ID);
};
