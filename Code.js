/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */

// GLOBAL VARIABLES - TODO : Set this variables in ScriptProperties !
var DIALOG_EVENT_TITLE = 'Novo Evento';
var DIALOG_CASHFLOW_TITLE = 'Registrar Evento';
var SHEET_EVENT_NAME = '_Eventos';

  // Title variables
var t_total_income = "Cachê";
var t_partial_income = "Recebido";
var t_net_result = "Líquido";
var t_leftover = "Resto";

var t_expenses = "Gastos";
var t_fees = "Cachê";
var t_total = "TOTAL";
var t_received = "Recebido";

var t_production = "Produção";
var t_cashflow = "Caixa";

  // Cells style
var s_money = '[$R$ -416]#,##0';
var s_percent = '0%';
var s_first_bg = "#fff3f9";
var s_extra_fee_bg = "#fef8e3";
var s_total_bg = "#eff4f8";
var s_second_font = "#999999";
var s_validation_bg = '#aefab2';

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Novo Evento', 'newEventDialog')
      .addItem('Registrar Evento', 'recordCashflowDialog')
      .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

/**
* Opens a dialog to create New Event. The dialog structure is described in the NewEvent.html
* project file.
*/
function newEventDialog() {
   var ui = HtmlService.createTemplateFromFile('NewEvent')
       .evaluate()
       .setWidth(350)
       .setHeight(500)
       .setSandboxMode(HtmlService.SandboxMode.IFRAME);
   SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_EVENT_TITLE);
}
 
 
/**
  * Opens a dialog. The dialog structure is described in the RecordCashflow.html
  * project file.
  */
function recordCashflowDialog() {
  var ui = HtmlService.createTemplateFromFile('RecordCashflow')
  .evaluate()
  .setWidth(350)
  .setHeight(200)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_CASHFLOW_TITLE);
}

// Return an array of the Sheets names that can receive CashFlow
// from an event.
function getCashflowNames() {
  var cashflowNames = [];
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  
  for (var i=0 ; i<sheets.length ; i++) {
    var sheet_name = sheets[i].getName();
    if (sheet_name.indexOf("_") != 0) {
      cashflowNames.push(sheet_name);
    }
  }
  return cashflowNames 
}

// Return an array of the event names on the Event sheet
function getEventNames() {
  var eventNames = [];
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet_events = spreadsheet.getSheetByName(SHEET_EVENT_NAME);
  var first_row = sheet_events.getRange(1,1,1,200).getValues()[0];
  
  for (var i=0 ; i<first_row.length ; i++) {
    var cel = first_row[i];
    var prohibited_val = ["", t_production, t_cashflow];
    
    if (typeof cel == 'string' && prohibited_val.indexOf(cel) == -1) {
      eventNames.push(cel);
    }
  }
  return eventNames 
}


/*****************************************************************************************
 * Record the money received by each member in his own sheet
 *
 * @param {string} event_name : the event's name selected in the Dialog
 *
 */
function recordEvent(event_name) {

  var spreadsheet = SpreadsheetApp.getActive();
  var sheet_events = spreadsheet.getSheetByName(SHEET_EVENT_NAME);
  var first_row = sheet_events.getRange(1,1,1,200).getValues()[0];
  
  // Catch the event column
  for (var i=0 ; i<first_row.length ; i++) {
    if (first_row[i] == event_name) {
      var event_col = i + 1;
      break;
    }
  }
  // Catch the last column of the event range
  var event_range_large = sheet_events.getRange(4, event_col - 1, 40, 20).getValues();
  for (var j = 0; j < event_range_large[0].length; j++) {
    if (event_range_large[0][j] == t_received) {
      var last_col = j;
      break;
    }
  }
  // Catch the perfect event range using the member's list that can receive cashflow
//  var members = getCashflowNames();
  
  var last_row = 1;
  for (var i = 1; i < event_range_large.length; i++) {
    if (event_range_large[i][0] != "") {
      last_row = i;
    }
  }
  
  var event_range = sheet_events
    .getRange(4, event_col - 1, last_row + 1,last_col + 1)
    .getValues();
 
  // Create a cashDivision object with organized and selected values to record
  var cashDivision = {};
  
  for (var i = 1; i < event_range.length; i++) {
    cashDivision[event_range[i][0]] = {};
    
    for (var j = 1; j < event_range[0].length; j++) {
      if (event_range[0][j] == t_fees) {
        cashDivision[event_range[i][0]][t_fees] = event_range[i][j];
        
      } else if (event_range[0][j] == t_production) {
          cashDivision[event_range[i][0]][t_production] = event_range[i][j];
      
      } else if (event_range[0][j] == t_cashflow) {
          cashDivision[event_range[i][0]][t_cashflow] = event_range[i][j];
      }
    }
  }
  
  // Special function to record the incomes data stored in "cashDivision[name]" object
  // into the "sheets[i]" sheet
  function recordInSheet(persIncomes, Sheet) {
    var first_row = Sheet.getRange(1,1,1, 20).getValues()[0];
    
    for (var title in persIncomes) {
    
      if (persIncomes[title] != "") {
      
        for (var j = 0; j < first_row.length; j++) {
          if (first_row[j].toLowerCase().indexOf(title.toLowerCase()) == 0) {
          
            if (title.toLowerCase().indexOf(t_cashflow.toLowerCase()) == 0) {
              // ****FILL CASHFLOW INCOME****
              var col_income = j + 4;
              var col_date = j + 6;
              var col_name = j + 5;
              var row_insert = 4;
              var name_income = title + " " + event_name;
              
              // Check if the event record doesn't exist already
              var actual_record = Sheet.getRange(4, j + 4, 30, 3).getValues();
              var isnew_record = true;
              
              
              
              for (var i = actual_record.length - 1; i >= 0 ; i--) {
              Logger.log("Nome record" + actual_record[i][col_name - j - 1 - 3]);
                if (actual_record[i][col_name - j - 4] == name_income) {
                  row_insert = 4 + i;
                  isnew_record = false;
                }
              }
              // Insert New row if is a new record
              if (isnew_record) {
                Sheet.getRange(row_insert, col_income, 1, 3).insertCells(SpreadsheetApp.Dimension.ROWS);
              }
              // Fill the cells
              Sheet.getRange(row_insert, col_income).setValue(persIncomes[title]);
              Sheet.getRange(row_insert, col_date).setValue(new Date());
              Sheet.getRange(row_insert, col_name).setValue(title + " " + event_name);
            
            } else {
            // ****FILL OTHER INCOMES****
              var col_income = j + 1;
              var col_date = j + 2;
              var col_name = j + 3;
              var row_insert = 4;
              var name_income = title + " " + event_name;
              
              // Check if the event record doesn't exist already
              var actual_record = Sheet.getRange(4, j + 1, 30, 3).getValues();
              var isnew_record = true;
              
              for (var i = actual_record.length - 1; i >= 0 ; i--) {
                if (actual_record[i][col_name - j - 1] == name_income) {
                  row_insert = 4 + i;
                  isnew_record = false;
                }
              }
              // Insert New row if is a new record
              if (isnew_record) {
                Sheet.getRange(row_insert, col_income, 1, 3).insertCells(SpreadsheetApp.Dimension.ROWS);
              }
              // Fill the cells
              Sheet.getRange(row_insert, col_income).setValue(persIncomes[title]);
              Sheet.getRange(row_insert, col_date).setValue(new Date());
              Sheet.getRange(row_insert, col_name).setValue(title + " " + event_name);
            }
          }
        }
      }
    }
  }
  
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  
  for (var i=0 ; i<sheets.length ; i++) {
    var sheet_name = sheets[i].getName();
    
    Logger.log("*******" + sheet_name + "*********");
    
    for (var name in cashDivision) {
      if (sheet_name == name) {
        // record the incomes data stored in the cashDivision[name] object into sheets[i]
        recordInSheet(cashDivision[name], sheets[i]);
        continue;
      }
    }
  }
}


/*****************************************************************************************
 * Create the formatted columns of a new account event
 *
 * Exemples of parameters received by the client-side JS : 
 * 
 * var name_event = "Nome do Evento";
 * var date_event = "01/01/2020";
 * var income_event = 5000;
 * var members_list = ["Pano", "Michel", "Daniel", "Clément", "Conta INTER", "Tutuka"];
 *
 *  var extra_fee = [
 *   { production : true},
 *   { cashflow : true}
 * ];
 *
 */
function createNewEvent(
  name_event,
  date_event,
  income_event,
  members_list,
  extra_fee
) {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet_events = spreadsheet.getSheetByName(SHEET_EVENT_NAME);

  var partial_income = 0;
  var t_income = [t_total_income, t_partial_income, t_net_result, t_leftover];
  var t_perperson = [t_expenses, t_fees, t_total, t_received];
  
  // Define a better list of objects to build extra_fee columns
  var obj_extra_fee = [  
    {
      istrue : false,
      title : t_production,
      percent : "10%",      
    },
    {
      istrue : false,
      title : t_cashflow,
      percent : "20%",      
    }
  ];
  
  // Adapt obj_extra_fee to the parameter extra_fee's values
  if (extra_fee[0]["production"]) {
    obj_extra_fee[0]["istrue"] = true
  }
  
  if (extra_fee[1]["cashflow"]) {
    obj_extra_fee[1]["istrue"] = true
  }  
  
  // Define the additional columns number due to the number
  // of extra_fee's checkbox checked
  var add_col = 0;
  
  for (var i = 0; i < obj_extra_fee.length; i++) {
    if (obj_extra_fee[i]["istrue"]) {
      add_col += 1;
    }
  }
  
  // Number of Event's columns
  var col_nb = 5 + add_col;

  // Insert Columns at the begining left
  sheet_events.insertColumnsBefore(1, col_nb);
  sheet_events.getRange(1, 1, 200, col_nb)
    .clearFormat()
    .clearDataValidations();

  // Format columns width
  sheet_events.setColumnWidth(1, 85);
  for (var i = 2; i <= col_nb; i++) {
    sheet_events.setColumnWidth(i, 65);
  }

// Fill text
  // Event
  sheet_events.getRange(1, 1)
    .setValue(date_event)
    .setHorizontalAlignment('right')
    .setNumberFormat('dd/MM/yyyy');

  sheet_events.getRange(1, 2)
    .setValue(name_event)
    .setHorizontalAlignment('left')
    .setFontWeight('bold');

  // Members
  for (var i = 0; i < members_list.length; i++) {
    sheet_events.getRange(5 + i, 1)
      .setValue(members_list[i])
      .setHorizontalAlignment('right')
      .setFontWeight('bold');
  }

  // Extra fee  
  
  var k = 0;
  for (var i = 0; i < obj_extra_fee.length; i++) {
    if (obj_extra_fee[i]["istrue"]) {
      k += 1;
      sheet_events.getRange(1, 5 + k)
        .setValue(obj_extra_fee[i]["title"])
        .setHorizontalAlignment('center');
        
      sheet_events.getRange(4, 3 + k)
        .setValue(obj_extra_fee[i]["title"])
        .setFontStyle("italic")
        .setFontColor(s_second_font)
        .setHorizontalAlignment('center');
       
      sheet_events.getRange(2, 5 + k)
        .setValue(obj_extra_fee[i]["percent"])
        .setHorizontalAlignment('center')
        .setNumberFormat(s_percent);
        
      var col_letter_person = String.fromCharCode(67 + k);
      var col_letter_percent = String.fromCharCode(69 + k);
      var formula_leftover = "=" + col_letter_percent + "2*C3-SUM(" + col_letter_person + "5:" + col_letter_person +"200)";
        
      sheet_events.getRange(3, 5 + k)
        .setFormula(formula_leftover)  // Leftover in the first extra_fee column
        .setHorizontalAlignment('center')
        .setBackground(s_extra_fee_bg)
        .setFontColor(s_second_font)
        .setNumberFormat(s_money);        
    }
  }

  // Income
  for (var i = 0; i < t_income.length; i++) {
    sheet_events.getRange(2, 1 + i)
      .setValue(t_income[i])
      .setHorizontalAlignment('center');
  }

  sheet_events.getRange(3, 1)
    .setValue(income_event)
    .setHorizontalAlignment('center')
    .setBackground(s_first_bg)
    .setNumberFormat(s_money);

  sheet_events.getRange(3, 2)
    .setValue(partial_income)
    .setHorizontalAlignment('center')
    .setBackground(s_first_bg)
    .setNumberFormat(s_money);

  sheet_events.getRange(2, 3)
    .setFontColor(s_second_font);

  sheet_events.getRange(3, 3)
    .setFormula('=A3-SUM(B5:B200)')
    .setHorizontalAlignment('center')
    .setBackground(s_first_bg)
    .setFontColor(s_second_font)
    .setNumberFormat(s_money);

  sheet_events.getRange(2, 4)
    .setFontColor(s_second_font);

  sheet_events.getRange(3, 4)
    .setFormula('=C3-SUM(F5:F200)')
    .setHorizontalAlignment('center')
    .setBackground(s_first_bg)
    .setFontColor(s_second_font)
    .setNumberFormat(s_money);

  // Per person
  sheet_events.getRange(4, 2)
    .setValue(t_expenses)
    .setHorizontalAlignment('center')
    .setFontStyle("italic")
    .setFontColor(s_second_font);

  sheet_events.getRange(4, 3)
    .setValue(t_fees)
    .setHorizontalAlignment('center')
    .setFontStyle("italic")
    .setFontColor(s_second_font);

  sheet_events.getRange(4, 4 + add_col)
    .setValue(t_total)
    .setHorizontalAlignment('center')
    .setFontStyle("italic")
    .setFontColor(s_second_font);

  sheet_events.getRange(4, 4 + add_col + 1)
    .setValue(t_received)
    .setHorizontalAlignment('center')
    .setFontStyle("italic");

  // Set alternative color and money format on personal lines
  
    // On personal income columns
  var r_income = sheet_events.getRange(5, 2, members_list.length + 1, 2);
  
  r_income.setNumberFormat(s_money)
    .setHorizontalAlignment('center');
  
  r_income.applyRowBanding()
    .setHeaderRowColor(null)
    .setFirstRowColor(s_first_bg)
    .setSecondRowColor('#ffffff')
    .setFooterRowColor(null);
  
    // On extra_fee columns
  
  if (add_col > 0) {
    var r_extra_fee = sheet_events.getRange(5, 4, members_list.length + 1, add_col);
    
    r_extra_fee.setNumberFormat(s_money)
      .setHorizontalAlignment('center');
    
    r_extra_fee.applyRowBanding()
      .setHeaderRowColor(null)
      .setFirstRowColor(s_extra_fee_bg)
      .setSecondRowColor('#ffffff')
      .setFooterRowColor(null);
  }
  
    // On TOTAL-Received columns
  var r_total_received = sheet_events.getRange(5, 4 + add_col, members_list.length + 1, 2);
  
  r_total_received.setNumberFormat(s_money)
    .setHorizontalAlignment('center');
  
  r_total_received.applyRowBanding()
    .setHeaderRowColor(null)
    .setFirstRowColor(s_total_bg)
    .setSecondRowColor('#ffffff')
    .setFooterRowColor(null);

  // Set the font color and the TOTAL formula on the TOTAL column 
  
  var first_total_cell = sheet_events.getRange(5,4 + add_col);
  var r_total = sheet_events.getRange(5, 4 + add_col, members_list.length, 1);  
  var formula_total = "=SUM(B5:" + String.fromCharCode(67 + add_col) + "5)";
  
  first_total_cell.setFormula(formula_total)
    .copyTo(r_total);
  
  r_total.setFontColor(s_second_font);

  // Set Borders
  sheet_events.getRange(1, 1, 200, col_nb)
    .setBorder(null, true, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet_events.getRange(1, 1, 3, col_nb)
    .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  sheet_events.getRange(5, 1, members_list.length + 3, col_nb)
    .setBorder(null, null, null, null, true, null, '#000000', SpreadsheetApp.BorderStyle.DOTTED);
    
  sheet_events.getRange(4, 1, 1, col_nb)
    .setBorder(true, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    
  sheet_events.getRange(2, 1, 2, 4)
    .setBorder(true, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    
    //Set Borders to the extra_fee infos
  if (add_col > 0) {
    sheet_events.getRange(1, 6, 1, add_col)
      .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
      
    sheet_events.getRange(1, 6, 3, add_col)
      .setBorder(null, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
      
    sheet_events.getRange(4, 4, 1, add_col)
      .setBorder(null, true, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.DOTTED);
  }
  
  sheet_events.getRange(4, 3)
    .setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.DOTTED);
    
    
  // Set Protection on cells with formulas
  sheet_events.getRange(3, 3, 1, 3)
    .protect()
    .setDescription('Protect formula changes on Incomes cells')
    .setWarningOnly(true);
    
  sheet_events.getRange(5, 4 + add_col, members_list.length + 1, 1)
    .protect()
    .setDescription('Protect formula changes on TOTAL cells')
    .setWarningOnly(true);
    
  if (add_col > 0) {
    sheet_events.getRange(1, 6, 3, add_col)
      .protect()
      .setDescription('Protect formula changes on Extra_fee cells')
      .setWarningOnly(true);  
  }
  
  // Set conditional formatting on TOTAL-Received columns
  
  var cond_formula = '=$'.concat(
    String.fromCharCode(69 + add_col),
    '5>=ROUNDDOWN($',
    String.fromCharCode(68 + add_col),
    '5)'
  );
  
  var conditionalFormatRules = sheet_events.getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([sheet_events.getRange(5,4 + add_col, members_list.length, 2)])
    .whenFormulaSatisfied(cond_formula)
    .setBackground(s_validation_bg)
    .build());
    
     // Conditional Formatting also for Income-received
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([sheet_events.getRange(3, 1, 1, 2)])
    .whenFormulaSatisfied("=$B3>=ROUNDDOWN($A3)")
    .setBackground(s_validation_bg)
    .build());  
    
  sheet_events.setConditionalFormatRules(conditionalFormatRules);
  
  // Prevent writing non-numbers in personal cells
  sheet_events.getRange(5, 2, members_list.length, 2 +  add_col).setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .setHelpText('Escreve um número')
    .requireFormulaSatisfied('=ISNUMBER(B5)')
    .build());
    
  sheet_events.getRange(3, 2).setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .setHelpText('Escreve um número')
    .requireFormulaSatisfied('=ISNUMBER(B3)')
    .build());
    
   sheet_events.getRange(5, 5 + add_col, members_list.length, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .setHelpText('Escreve um número')
    .requireFormulaSatisfied('=ISNUMBER(' + String.fromCharCode(69 + add_col) + '5)')
    .build()); 
        
  
  // Set Active cell on the Event Name
  var name_event_cell = sheet_events.getRange(1,2);
  sheet_events.setActiveRange(name_event_cell);
}
