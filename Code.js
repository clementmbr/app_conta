/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */

var DIALOG_TITLE = 'Example Dialog';
var SIDEBAR_EVENT_TITLE = 'Novo Evento';

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Novo Evento', 'newEventSidebar')
      // .addItem('Show dialog', 'showDialog')
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
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function newEventSidebar() {
  var ui = HtmlService.createTemplateFromFile('SidebarNewEvent')
      .evaluate()
      .setTitle(SIDEBAR_EVENT_TITLE)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui); // "showSidebar" is a special SpreadsheetApp funcion
}

// /**
//  * Opens a dialog. The dialog structure is described in the Dialog.html
//  * project file.
//  */
// function showDialog() {
//   var ui = HtmlService.createTemplateFromFile('Dialog')
//       .evaluate()
//       .setWidth(400)
//       .setHeight(190)
//       .setSandboxMode(HtmlService.SandboxMode.IFRAME);
//   SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_TITLE);
// }

/**
 * Create the formatted columns of a new account event
 *
 * Exemples of parameters : 
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
  // Use data collected from sidebar to manipulate the sheet.

  var spreadsheet = SpreadsheetApp.getActive();
  var sheet_events = spreadsheet.getSheetByName("Eventos");

  var partial_income = 0;
  
  // Define a better list of objects to build extra_fee columns
  var obj_extra_fee = [  
    {
      istrue : false,
      title : "Produção",
      percent : "10%",      
    },
    {
      istrue : false,
      title : "Caixa",
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
  
  // Define title variables
  
  var t_total_income = "Cachê";
  var t_partial_income = "Recebido";
  var t_net_result = "Líquido";
  var t_leftover = "Resto";
  var t_income = [t_total_income, t_partial_income, t_net_result, t_leftover];

  var t_expenses = "Gastos";
  var t_fees = "Cachês";
  var t_total = "TOTAL";
  var t_received = "Recebido";
  var t_perperson = [t_expenses, t_fees, t_total, t_received];

  // Cells style
  var s_money = '[$R$ -416]#,##0';
  var s_percent = '0%';
  var s_first_bg = "#fff3f9";
  var s_extra_fee_bg = "#fef8e3";
  var s_total_bg = "#eff4f8";
  var s_second_font = "#999999";

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
      var formula_leftover = "=" + col_letter_percent + "2*(C3-SUM(" + col_letter_person + "5:" + col_letter_person +"200))";
        
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
  sheet_events.getRange(1, 1, 200, col_nb).setBorder(null, true, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet_events.getRange(1, 1, 3, col_nb).setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}





