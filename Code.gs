 /*
  *  Peer-Grading Tool for Google Spreadsheets
  *  Copyright 2018 Stefan Stolz and Nina Margreiter

  *  Licensed under the Apache License, Version 2.0 (the "License");
  *  you may not use this file except in compliance with the License.
  *  You may obtain a copy of the License at

  *  http://www.apache.org/licenses/LICENSE-2.0

  *  Unless required by applicable law or agreed to in writing, software
  *  distributed under the License is distributed on an "AS IS" BASIS,
  *  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
  *  See the License for the specific language governing permissions and
  *  limitations under the License.
  */

// #### Todo ####
// TODO: settings['colLastGrade'] fertig machen
// TODO: remove bedingte Formatierung from whole sheet before setting it if something has changed (newDataValidation())

var version = '1.0';
var manualLink = 'https://github.com/ststolz/Peer-Grading-Tool/blob/master/README.md';

var settings = new Object();
settings['version'] = version;
settings['colFirstGrade'] = "C";
settings['colLastGrade'] = "G";
settings['spalteUsernamen'] = "B";
settings['rowFirstData'] = "4"; 
settings['punkteMin'] = "0";
settings['punkteMax'] = "2";
settings['noten'] = false;
settings['nameVorlage'] = "Template";
settings['deletionMarker'] = "*"; 
settings['sheetAuswertungName'] = "Evaluation"; 
settings['maxPointsText'] = "Weighted Points";
settings['endRatingText'] = "Percent";
settings['ratedPersonText'] = "Rated Student";
settings['averageRatingText'] = "Average";
settings['medianRatingText'] = "Median";
settings['competencesText'] = "Competences";

var documentProperties = PropertiesService.getDocumentProperties();

var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var spreadsheets=activeSpreadsheet.getSheets();

var vorlageSheet;
var vorlageSheetRange;
var vorlageSheetData;
// User Array
var users;
//Feststellen wie viele Spalten und Zeilen in Tabelle  
var rowsN;
var colsN;

getDocumentProperties();

function initTemplateSheet(){
  vorlageSheet = searchSheet(settings['nameVorlage']);
  // ### Search for Vorlage ###  
  if( ! vorlageSheet){
    Browser.msgBox('No Sheet with name '+settings['nameVorlage']+' found. I am loading an example Template. Please adapt this and start again.');
    sheetFromTemplate("1NjQnLFDq6FQ4eXUt9lQOfMj8OD_jQB169i9JbPAcH08", settings['nameVorlage']);
    throw new Error("Stopping execution - no template found!");
  }  
  vorlageSheetRange = vorlageSheet.getDataRange();
  vorlageSheetData = vorlageSheetRange.getValues();
  rowsN = vorlageSheetData.length;
  colsN = vorlageSheetData[0].length;
}

function getDocumentProperties(){
	for (var k in settings){
	    if (typeof settings[k] !== 'function') {
	    	Logger.log("Key is " + k + ", value is: " + settings[k]);
	    	var property = documentProperties.getProperty(k);
	    	if(property){
	    		Logger.log("Saved Property is: " + property);	    		
	    		settings[k]=property;
	    	}
	    	else {
	    		Logger.log("Not Saved Property is: " + k + ", saving: "+settings[k]);
	    		documentProperties.setProperty(k,settings[k]);
	    	}
	    }	    
	}
}

function saveDocumentProperties(saveSettings){
	for (var k in settings){
		documentProperties.setProperty(k,saveSettings[k]);
	}
}

function onOpen() {
  
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Vortragstools');        
  
  menu
    .addItem('Show sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Manual', 'showManual'); 
  
  menu.addItem('Version '+version, 'disclaimer')
  .addToUi();
  
}

function showAnchor(title,name,url) {
  var html = '<html><body><a href="'+url+'" target="blank" onclick="google.script.host.close()">'+name+'</a></body></html>';
  var ui = HtmlService.createHtmlOutput(html)
  SpreadsheetApp.getUi().showModelessDialog(ui,title);
}
  
function showManual(){
	showAnchor('Manual','Link to Peer-Grading Tool Manual', manualLink);
}

function getSettings(){
	var returnSettings = new Object();
	for (var k in settings){
		returnSettings[k] = settings[k].toString();
	}
	return returnSettings;
}

function showSidebar() {
	  var html = HtmlService.createHtmlOutputFromFile('sidebar.html')
	      .setTitle('Peer Grading Tool v'+version)
	      .setWidth(300);
	  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
	      .showSidebar(html);
}

function disclaimer(){
  Browser.msgBox('Version '+version+' of Peer-Grading Tool. See manual for license.');
}

function calcMean(){
}

function searchSheet(name){
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheets=activeSpreadsheet.getSheets();
  for (var i = 0; i < spreadsheets.length; i++) {
    if(spreadsheets[i].getName() == name){
      return spreadsheets[i];
    }
  }
  return false;
}

function initUsernameArr(){
  users = generateUsernameArr();
}

function generateUsernameArr(){
  Logger.log("Generate Usernames!");
  
  var users = new Array(vorlageSheetData.length-(settings['rowFirstData']-1));
  
  for (var k = (settings['rowFirstData']-1), i = 0; k < vorlageSheetData.length; k++, i++) {      
    var username = trim(vorlageSheetData[k][letterToColumn(settings['spalteUsernamen'])-1]);
    users[i] = username;
    Logger.log("generate: "+username);
  }
  return users;  
}

function getTeamplateUserField(username){
  userCell = null;
  for (var k = (settings['rowFirstData']-1), i = 0; k < vorlageSheetData.length; k++, i++) {      
    if(username == trim(vorlageSheetData[k][letterToColumn(settings['spalteUsernamen'])-1])){
      var userCell = vorlageSheet.getRange((k+1), letterToColumn(settings['spalteUsernamen']));
    }    
  }
  return userCell;
}

function compareUsers(userArr){
  
}

function sheetFromTemplate(templateID, sheetName){
  var sheet = SpreadsheetApp.openById(templateID).getSheetByName(sheetName);
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var templateCopy = sheet.copyTo(activeSpreadsheet);
  insertedTemplate = activeSpreadsheet.insertSheet(sheetName, 2, {template: templateCopy}); 
  activeSpreadsheet.deleteSheet(templateCopy);
  insertedTemplate.deleteRow(1);
  return;
}

function processTemplateSheet(){
  if(!isOwner()) throw new Error("Stopping execution - you are not the owner of this spreadsheet!");
  
  // ### Protect Vorlage ###
  var protection = vorlageSheet.protect().setDescription('Template Protection');  
  var me = Session.getEffectiveUser();   
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
     
  /**
    * Validation that data in Range of allowed points
    */
  var vorlageSheetGradeRange = vorlageSheet.getRange(settings['rowFirstData'],letterToColumn(settings['colFirstGrade']),users.length,vorlageSheetRange.getLastColumn()-letterToColumn(settings['colFirstGrade']));
  var rule = SpreadsheetApp.newDataValidation().requireNumberBetween(parseFloat(settings['punkteMin']), parseFloat(settings['punkteMax'])).setAllowInvalid(false).build();
  vorlageSheetGradeRange.setDataValidation(rule);
  
  Logger.log("Setting edit rights for users");
  // ### Document Edit Rights for Users ###
  for(var i = 0; i < users.length; i++){ 
    var username = users[i];
    
    if(username != ""){
      activeSpreadsheet.addEditor(username);
      Logger.log(username);
    }
  }
  Logger.log("Edit rights set");
  
}

function startUpdateActiveUser(){
  Logger.log("Starting Acitve User Update");
  
  if(!isOwner()) throw new Error("Stopping execution - you are not the owner of this spreadsheet!");
  
  initTemplateSheet();
  initUsernameArr();
  
  var activeSheet=activeSpreadsheet.getActiveSheet();
  var username = activeSheet.getName();    
  
  if(username != settings['sheetAuswertungName'] && username.indexOf(settings['deletionMarker']) != 0 && username != settings['nameVorlage'] ){
    generateUser(true, username);
  }
}

function startGenerateAllUsers(update){
  
  Logger.log("Generating sheets for all users");
  
  if(!isOwner()) throw new Error("Stopping execution - you are not the owner of this spreadsheet!");
  
  initTemplateSheet();
  initUsernameArr();   
  Logger.log("Processing Template Sheet");
  processTemplateSheet();  
  
  Logger.log("Update user row to remove signs");
  if(update) {
    for (var r = settings['rowFirstData']; r <= rowsN; r++) {     
      Logger.log("Row: "+r);
      vorlageSheet.getRange(r, letterToColumn(settings['spalteUsernamen'])).setNumberFormat('"☒" @');
    }
  }
  
  // ### Walk through lines with Usernames ###
  Logger.log("Walking users array");
  for(var i = 0; i < users.length; i++){     
    
    var username = users[i];       
    if(username == ""){ continue; }
    
    Logger.log("Processing User: "+username);
    
    generateUser(update, username);
    
  }
  
  markNotExistingUsers();
  
}

function startGenerateSelectedUser(){
  
  Logger.log("Generating selected user");
  
  if(!isOwner()) throw new Error("Stopping execution - you are not the owner of this spreadsheet!");
  
  initTemplateSheet();
  
  username = activeSpreadsheet.getActiveCell().getValue();
  activeSpreadsheet.addEditor(username);
  
  generateUser(false, username);
  
}

function markNotExistingUsers(deleteOld){
  
  //### Mark not existing Users ###
  for (var i = 0; i < spreadsheets.length; i++) {
    var sheet = spreadsheets[i];
    var isUserExisting = false;
    for (var k = (settings['rowFirstData']-1); k < vorlageSheetData.length; k++) {
      var username = trim(vorlageSheetData[k][letterToColumn(settings['spalteUsernamen'])-1]);
      if(username == sheet.getName()){
        isUserExisting = true;
      }        
    }    
    if(!isUserExisting && sheet.getName().indexOf(settings['deletionMarker']) != 0 && settings['nameVorlage'] != sheet.getName() && settings['sheetAuswertungName'] != sheet.getName()){
      var sheetDeletionMarkerName = settings['deletionMarker'] + " " + sheet.getName();
      if(searchSheet(sheetDeletionMarkerName)){
    	  activeSpreadsheet.deleteSheet(searchSheet(sheetDeletionMarkerName));
      }      
      sheet.setName(sheetDeletionMarkerName);
    }
    else if(sheet.getName().indexOf(settings['deletionMarker']) == 0){
      activeSpreadsheet.deleteSheet(sheet);
    }
  }
}

function generateUser(update, username){    
  
  if(!isOwner()) throw new Error("Stopping execution - you are not the owner of this spreadsheet!");
  
  templateUserField = getTeamplateUserField(username);
  templateUserField.setNumberFormat('"☒" @');
  
  var isFirstRun = true; 
  
  if(spreadsheets.length > 1){    
      isFirstRun = false;   
  }       
    
  var newSheet;  
  var isExisting = false;
  
  // ### Check if Sheet is already existing ###
  if(!isFirstRun){
    Logger.log("Check if already existing");
    newSheet = searchSheet(username);
    if(newSheet && update){
      Logger.log("Is existing and updating");
      isExisting = true;
      var existingSheetData = newSheet.getDataRange().getValues();
      
      // Update complete sheet except grading data
      for (var r = 0; r < rowsN; r++) { 
        for(var c = 0; c < colsN; c++){          
          if(!(r > settings['rowFirstData']-2 && c > letterToColumn(settings['colFirstGrade'])-2 && c < colsN-1)){        			
            if(true){
              newSheet.getRange(r+1, c+1).setValue(vorlageSheetData[r][c]);
            }
          }  
        }        	
      }
    } 
    else if(newSheet) {
      Logger.log("Is existing");
      templateUserField.setNumberFormat('"☑" @');
      return;
    }
  }
  
  if(isFirstRun || ! isExisting){
    Logger.log("Make copy because doesn't exist");
    newSheet = vorlageSheet.copyTo(activeSpreadsheet);
    for (var r = settings['rowFirstData']; r <= rowsN; r++) {     
      newSheet.getRange(r, letterToColumn(settings['spalteUsernamen'])).setNumberFormat('@');        
    }         
  }
  
  
  // ### Remove old Protections ###
  var protections = newSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var j = 0; j < protections.length; j++) {
    var protection = protections[j];
    protection.remove();      
  }
  var protections = newSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var j = 0; j < protections.length; j++) {
    var protection = protections[j];
    protection.remove();      
  }
  
  // ### Sheet Protection ###
  var protection = newSheet.protect().setDescription('Sheet Protection');
  
  // Bereich berechnen in dem Reviews stehen
  var unprotected = newSheet.getRange(settings['rowFirstData'],letterToColumn(settings['colFirstGrade']),rowsN - (settings['rowFirstData']-1), colsN - letterToColumn(settings['colFirstGrade']));
  
  // Bereich mit Reviews von Protection ausnehmen
  protection.setUnprotectedRanges([unprotected]);    
  
  var me = Session.getEffectiveUser();
  
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  
  // ### Range Protection ###
  var protectRange = unprotected.protect().setDescription('Range Protection'); 
  protectRange.addEditor(me);
  protectRange.removeEditors(protectRange.getEditors());
  if (protectRange.canDomainEdit()) {
    protectRange.setDomainEdit(false);
  }
  protectRange.addEditor(username);
  
  //protection = protection.addEditor(username);
  if(newSheet.getName() != username){
    newSheet.setName(username);
  }
    
  templateUserField.setNumberFormat('"☑" @');
  
}

function startGenerateEvaluation() {
  
  if(!isOwner()) throw new Error("Stopping execution - you are not the owner of this spreadsheet!");
  
  
  initTemplateSheet(); 
  initUsernameArr();
  
  Browser.msgBox('Trying to make evaluation section for all users. If this exceeds maximum execution time, create the rest of the sheets per individual user.');
      
  // Walk through sheets to get users  
  for (var i = 0; i < spreadsheets.length; i++) {
    
    var bewerteterSheet = spreadsheets[i];  
    var bewerteter = bewerteterSheet.getName();
    
    if(bewerteter != settings['sheetAuswertungName'] && bewerteter.indexOf(settings['deletionMarker']) != 0 && bewerteter != settings['nameVorlage'] ){
      
      SpreadsheetApp.setActiveSheet(bewerteterSheet);
      
      createUserEvaluation(bewerteterSheet);
      
      
    }
  }
  //resizeAllColumns(yourNewSheet, 2);
}

function startGenerateEvaluationActiveUser() {
  
  Logger.log("Starting Acitve User Eval Gen");
  
  if(!isOwner()) throw new Error("Stopping execution - you are not the owner of this spreadsheet!");
  
  initTemplateSheet();
  initUsernameArr();
  
  
  
  var activeSheet=activeSpreadsheet.getActiveSheet();
  var bewerteter = activeSheet.getName();  
  
  
  if(bewerteter != settings['sheetAuswertungName'] && bewerteter.indexOf(settings['deletionMarker']) != 0 && bewerteter != settings['nameVorlage'] ){
    createUserEvaluation(activeSheet);  
  }
  
}

function createUserEvaluation(bewerteterSheet){
  
  // Get important Data
  var nCompetences = (parseInt(letterToColumn(settings['colLastGrade'])) - parseInt(letterToColumn(settings['colFirstGrade'])))+1;
  var bewerteter = bewerteterSheet.getName();
  var firstRowOfUserData;
  
  //yourNewSheet.getRange(row, 1).setValue(bewerteter);
  var row = 1;
  row++;row++;
      
      /*
       * Get ranges
       */
      var sheetData = bewerteterSheet.getDataRange().getValues();
      var sheetDataRows = sheetData.length;
      var skillNameRange = bewerteterSheet.getRange(settings['colFirstGrade']+(settings['rowFirstData']-2)+":"+settings['colLastGrade']+(settings['rowFirstData']-2));
      var skillWeightRange = bewerteterSheet.getRange(settings['colFirstGrade']+(settings['rowFirstData']-1)+":"+settings['colLastGrade']+(settings['rowFirstData']-1));
      var userNameCol = bewerteterSheet.getRange(settings['spalteUsernamen']+settings['rowFirstData']);
      var sheetName = bewerteterSheet.getName();
      
      // go to row after existing data 
      row += users.length + parseInt(settings['rowFirstData']);      
      
      /*
       * Create Array with competences
       */      
      var competences = new Array(nCompetences);
      
      for (var j = 0; j < nCompetences; j++) {
        competences[j]= new Array(sheetData.length);
      }
      Logger.log("competences.length: "+competences.length);
      
      /*
       * Set Row Headings
       */
      var competencesHeader = bewerteterSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))-1);
      var weightingHeader = bewerteterSheet.getRange(row+1, parseFloat(letterToColumn(settings['colFirstGrade']))-1);       
      competencesHeader.clearFormat().setValue(settings['competencesText']).setFontWeight("bold").setBackground(skillNameRange.getBackground());
      weightingHeader.clearFormat().setValue(settings['maxPointsText']).setFontWeight("bold").setBackground(skillWeightRange.getBackground());
            
      // Draw heading borders
      bewerteterSheet.getRange(columnToLetter(letterToColumn(settings['colFirstGrade'])-1)+row+":"+settings['colLastGrade']+row).setBorder(true, true, true, true, false, false);
      bewerteterSheet.getRange(columnToLetter(letterToColumn(settings['colFirstGrade'])-1)+(row+1)+":"+settings['colLastGrade']+(row+1)).setBorder(true, true, true, true, false, false);
                       
      // Walk through headings
      var numCols = sheetData[0].length - (letterToColumn(settings['colFirstGrade']));
      
      var row_weighting = row+1;
      
      for (var k = 0; k < nCompetences; k++) {
        var heading = sheetData[settings['rowFirstData']-3][letterToColumn(settings['colFirstGrade'])-1+k];        
        var weighting = sheetData[settings['rowFirstData']-2][letterToColumn(settings['colFirstGrade'])-1+k];
        
        var headingCell = bewerteterSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+k);
        headingCell.setValue("='"+sheetName+"'!"+columnToLetter(letterToColumn(settings['colFirstGrade'])-1+1+k)+(settings['rowFirstData']-2)).setBackground(skillNameRange.getBackground());
        
        var weightingCell = bewerteterSheet.getRange(row_weighting, parseFloat(letterToColumn(settings['colFirstGrade']))+k);
        weightingCell.setValue("='"+sheetName+"'!"+columnToLetter(letterToColumn(settings['colFirstGrade'])-1+1+k)+(settings['rowFirstData']-1)+"*"+settings['punkteMax']).setBackground(skillWeightRange.getBackground());
        
      }
        
      
      // Set Text for endrating
      Logger.log("#1: ");
      var resultsHeader = bewerteterSheet.getRange(row+1, parseInt(letterToColumn(settings['colLastGrade']))+1);
      Logger.log("#2");
      resultsHeader.clearFormat().setValue(settings['endRatingText']).setFontWeight("bold");
      Logger.log("#3");
      
      var counterUsers = 0;
      
      // set first row for Evaluation      
      row++;row++;
      firstRowOfUserData = row;
      Logger.log("#4");
      
      // Walk through sheets to get data for user
      for (var j = 0; j < spreadsheets.length; j++) {
        
        Logger.log("#5");
        
        var sheet = spreadsheets[j];
        var bewerter = sheet.getName();
        
        Logger.log("Bewerteter: "+bewerter);
        
        
        if(bewerter != settings['sheetAuswertungName'] && bewerter != bewerteter && sheet.getName().indexOf(settings['deletionMarker']) != 0 && settings['nameVorlage'] != sheet.getName()){
          
          
          
          bewerteterSheet.getRange(row, 2).clearFormat().setValue(bewerter).setBackground(userNameCol.getBackground());
          
          var sheetData = sheet.getDataRange().getValues();           
          
          // Walk through reviews
          // Walk through lines
          Logger.log("users.length: "+users.length);
          
          for (var k = 0; k < sheetData.length; k++) {
            Logger.log("k: "+k+"; users+fr: "+(users.length + parseInt(settings['rowFirstData'])));
            // If line is for reviewed person
            if(sheetData[k][letterToColumn(settings['spalteUsernamen'])-1] == bewerteter && k < users.length + parseInt(settings['rowFirstData'])){
              Logger.log("get it!");
              // calc number of cols with data
              var colsWithData = sheetData[k].length - (letterToColumn(settings['colFirstGrade'])-1);
              // Walk through all cols
              
              Logger.log("nCompetences: "+nCompetences);
              for(var l = 0; l < nCompetences+2; l++){
                var val = sheetData[k][parseFloat(letterToColumn(settings['colFirstGrade']))+l-1];
                
                if(l<nCompetences){
                  // Fill Array for later mean calc
                  // Only if in range of points
                  if(isInPointsRange(val) && l < colsWithData -1){
                    competences[l][counterUsers] = val;                                 
                  }
                  else{
                    var cell = bewerteterSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+l).clearFormat().setValue("");
                  }
                  
                  var columnGrade = columnToLetter(parseFloat(letterToColumn(settings['colFirstGrade']))+l);
                  /* set link to user rating value */
                  var cell =  bewerteterSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+l).clearFormat().setValue("=IF('"+bewerter+"'!"+columnGrade+(k+1)+" <> \"\"; '"+bewerter+"'!"+columnGrade+(k+1)+" * '"+settings['nameVorlage']+"'!"+columnGrade+(settings['rowFirstData']-1)+";\"\")");
                }
                
                /* Calc Col Mean */
                Logger.log("col: "+l+"; lastgrade: "+letterToColumn(settings['colLastGrade']));
                if(l+letterToColumn(settings['colFirstGrade']) == letterToColumn(settings['colLastGrade'])+1){               	 
                  Logger.log("get it!");
                  var rangeGrades = settings['colFirstGrade']+(row)+":"+settings['colLastGrade']+(row); 
                  var rangeWeights = settings['colFirstGrade']+(row_weighting)+":"+settings['colLastGrade']+(row_weighting);
                  
                  var sumWeights = "SUMIFS("+rangeWeights+";'"+bewerter+"'!"+settings['colFirstGrade']+(k+1)+":"+settings['colLastGrade']+(k+1)+";\"<>\")";
                  bewerteterSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+l).clearFormat().setValue("=IF("+sumWeights+">0;ROUND(SUM("+rangeGrades+")/("+sumWeights+");2);\"\")").setNumberFormat("#.###%");
                  
                }       
                
              }               
            }  
          }         
          
          row++;
          counterUsers++;
        }
      }
      //draw border around ratings
      bewerteterSheet.getRange(columnToLetter(letterToColumn(settings['colFirstGrade'])-1)+(row-counterUsers)+":"+settings['colLastGrade']+(row-1)).setBorder(true, true, true, true, false, false);
      // Calc mean values
      bewerteterSheet.getRange(row, letterToColumn(settings['colFirstGrade'])-1).clearFormat().setValue(settings['averageRatingText']).setFontWeight("bold");
      for (var x = 0; x < nCompetences+1; x++) {
        var calcCol = columnToLetter(parseFloat(letterToColumn(settings['colFirstGrade']))+x);
        var rangeForCalc = calcCol+firstRowOfUserData+":"+calcCol+(row-1);
        Logger.log("competences.length-1: "+competences.length+"; x: "+x);
        if(x == nCompetences){
          bewerteterSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+x).clearFormat().setValue("=IF(COUNT("+rangeForCalc+")>0;ROUND(AVERAGE("+rangeForCalc+");2);\"\")").setNumberFormat("#.###%");
        }
        else{
          bewerteterSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+x).clearFormat().setValue("=IF(COUNT("+rangeForCalc+")>0;ROUND(AVERAGE("+rangeForCalc+");1) &\" / \"& "+calcCol+(firstRowOfUserData-1)+";\"\")").setNumberFormat("0.0");
        }
      }
      row++;
      bewerteterSheet.getRange(row, letterToColumn(settings['colFirstGrade'])-1).clearFormat().setValue(settings['medianRatingText']).setFontWeight("bold");
      for (var x = 0; x < nCompetences+1; x++) {
        var calcCol = columnToLetter(parseFloat(letterToColumn(settings['colFirstGrade']))+x);
        var rangeForCalc = calcCol+firstRowOfUserData+":"+calcCol+(row-2);
        Logger.log("competences.length-1: "+competences.length+"; x: "+x);
        if(x == nCompetences){
          bewerteterSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+x).clearFormat().setValue("=IF(COUNT("+rangeForCalc+")>0;ROUND(MEDIAN("+rangeForCalc+");2);\"\")").setNumberFormat("#.###%");
        }
        else{
          bewerteterSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+x).clearFormat().setValue("=IF(COUNT("+rangeForCalc+")>0;ROUND(MEDIAN("+rangeForCalc+");1) &\" / \"&"+calcCol+(firstRowOfUserData-1)+";\"\")").setNumberFormat("0.0");
        }
      }
}


function resizeAllColumns(sheet,num){
  
  var sheetData = sheet.getDataRange().getValues();  
           
  
  // Walk through all cols
  for(var l = 1; l <= num; l++){
    Logger.log("Col: "+l);
    sheet.autoResizeColumn(l);
    
  }
}

function isInPointsRange(val){
  if(isNumeric(val) && val >= settings['punkteMin'] && val <= settings['punkteMax']){
    return true;
  }
  return false;
}

function isNumeric(s) {
  return !isNaN(parseFloat(s)) && isFinite(s);
}

function trim(stringToTrim) {
	return stringToTrim.replace(/^\s+|\s+$/g,"");
}

function isOwner(){
	

	return true;
	
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();  
  
  var owner = activeSpreadsheet.getOwner();  
  var activeUser =  Session.getActiveUser();
  
  Logger.log('activeUser id: '+ activeUser.getEmail() + ' owner id: ' + owner.getEmail());
  
  if(owner.getEmail() == activeUser.getEmail()){ 
	  Logger.log('Is owner!');
	  return true; }
  else { 
	  Logger.log('Not owner!');
	  return false};
}

function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}