 /*
  *  Peer-Grading Tool for Google Spreadsheets
  *  Copyright (C) 2018  Stefan Stolz and Nina Margreiter

  *  This program is free software: you can redistribute it and/or modify
  *  it under the terms of the GNU General Public License as published by
  *  the Free Software Foundation, either version 3 of the License, or
  *  (at your option) any later version.

  *  This program is distributed in the hope that it will be useful,
  *  but WITHOUT ANY WARRANTY; without even the implied warranty of
  *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  *  GNU General Public License for more details.

  *  You should have received a copy of the GNU General Public License
  *  along with this program.  If not, see <http://www.gnu.org/licenses/>.
  */

// #### Bedingungen an die Tabelle ###
// Die Tabelle muss nach der Auswertung noch genau eine frei wählbare Spalte haben (z.B. Durchschnitt)
// Unter den Kompetenz Überschriften muss eine Zeile mit Gewichtungen kommen

// #### Todo ####
// TODO: settings['colLastGrade'] fertig machen

var version = 0.50; 
var settings = new Object();
settings['version'] = version;
settings['colFirstGrade'] = "C";
settings['colLastGrade'] = "H";
settings['spalteUsernamen'] = "B";
settings['rowFirstData'] = 2; 
settings['punkteMin'] = 0;
settings['punkteMax'] = 2;
settings['noten'] = false;
settings['nameVorlage'] = "Vorlage";
settings['deletionMarker'] = "*"; 
settings['sheetAuswertungName'] = "Auswertung"; 
settings['maxPointsText'] = "Erreichbare Punkte";
settings['endRatingText'] = "Gesamtwertung";

var documentProperties = PropertiesService.getDocumentProperties();

getDocumentProperties();

function getDocumentProperties(){
	for (var k in settings){
	    if (typeof settings[k] !== 'function') {
	    	Logger.log("Key is " + k + ", value is: " + settings[k]);
	    	var property = documentProperties.getProperty(k);
	    	if(property){
	    		Logger.log("Saved Property is: " + property);
	    		switch(k){
	    		    case 'version':
	    			case 'noten': 
	    			case 'nameVorlage':
	    			case 'deletionMarker':
	    			case 'sheetAuswertungName':
	    			case 'maxPointsText':
	    			case 'endRatingText':
                    case 'colFirstGrade':
                    case 'colLastGrade':
                    case 'spalteUsernamen':
	    				property = property;
	    				break;
	    			default:
	    				property = parseFloat(property);
	    				break;
	    		} 
	    		settings[k]=property;
	    	}
	    	else {
	    		Logger.log("Not Saved Property is: " + k + ", saving: "+settings[k]);
	    		documentProperties.setProperty(k,settings[k]);
	    	}
	    }
	    
	}
	
	
	
	//documentProperties.getProperty();
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
  
function showManual(){
	var html = HtmlService.createHtmlOutputFromFile('help');
	  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
	      .showModalDialog(html, 'Manual'); 
}

function getSettings(){
	var returnSettings = new Object();
	for (var k in settings){
		returnSettings[k] = settings[k].toString();
	}
	return returnSettings;
}

function showSidebar() {
	  var html = HtmlService.createHtmlOutputFromFile('SIDEBAR3.html')
	      .setTitle('Peer Grading Tool v'+version)
	      .setWidth(300);
	  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
	      .showSidebar(html);
}

function disclaimer(){
  Browser.msgBox('Version '+version+' des Feedback Tools erstellt von MARG und STOL (HAK Imst)');
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

function generateUsernameArr(){
	Logger.log("Usernames!");
  var vorlageSheet = searchSheet(settings['nameVorlage']);
  var vorlageSheetData = vorlageSheet.getDataRange().getValues();
  var users = new Array(vorlageSheetData.length-(settings['rowFirstData']-1));
  
  for (var k = (settings['rowFirstData']-1), i = 0; k < vorlageSheetData.length; k++, i++) {      
    var username = trim(vorlageSheetData[k][letterToColumn(settings['spalteUsernamen'])-1]);
    users[i] = username;
    Logger.log("generate: "+username);
  }
  return users;  
}

function compareUsers(userArr){
  
}


function generateSheets(){    
  
  Logger.log("Generate Sheets");
  Logger.log("sheetAuswertungName: "+settings['sheetAuswertungName']);
	
  if(!isOwner()) return;  
  
  var isFirstRun = true;
  
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheets=activeSpreadsheet.getSheets();
  
  var ui = SpreadsheetApp.getUi(); 
  if(spreadsheets.length > 1){
    //var result = ui.alert(
    //  'Achtung!',
    //  'Bei einer Änderung eines Usernamens werden die Daten des alten Usernamens gelöscht und ein neues Blatt erzeugt! Trotzdem fortfahren?',
    //  ui.ButtonSet.YES_NO);
    
    // Process the user's response.
    //if (result == ui.Button.YES) {
      isFirstRun = false;
    //} 
    //else {
    //  return;
    //}
    
  }
  
  // ### Search for Vorlage ###
  var vorlageSheet = searchSheet(settings['nameVorlage']);
  if( ! vorlageSheet){
    Browser.msgBox('Fehler! Kein Sheet mit dem Namen '+settings['nameVorlage']+' gefunden. Breche ab...');
    return;
  }
  
  
  
  var vorlageSheetRange = vorlageSheet.getDataRange();
  
  
  // ### Protect Vorlage ###
  var protection = vorlageSheet.protect().setDescription('Sheet Protection');  
  var me = Session.getEffectiveUser();   
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  
  var vorlageSheetData = vorlageSheetRange.getValues();
  //Feststellen wie viele Spalten und Zeilen in Tabelle
  // * var sheetData = vorlageSheet.getDataRange().getValues();    
  var rowsN = vorlageSheetData.length;
  var colsN = vorlageSheetData[0].length;
  
  Logger.log("Start username ");
 
  // Get User Array
  var users = generateUsernameArr();
  //### Validation that data in Range of allowed points
  var vorlageSheetGradeRange = vorlageSheet.getRange(settings['rowFirstData'],letterToColumn(settings['colFirstGrade']),users.length,vorlageSheetRange.getLastColumn()-letterToColumn(settings['colFirstGrade']));
  var rule = SpreadsheetApp.newDataValidation().requireNumberBetween(parseFloat(settings['punkteMin']), parseFloat(settings['punkteMax'])).setAllowInvalid(false).build();
  vorlageSheetGradeRange.setDataValidation(rule);
  
    
  // ### Document Edit Rights for Users ###
  for(var i = 0; i < users.length; i++){ 
    var username = users[i];
    
    if(username != ""){
      activeSpreadsheet.addEditor(username);
      Logger.log(username);
    }
  }
  
  
  // ### Walk through lines with Usernames ###
  for(var i = 0; i < users.length; i++){ 
    
    var username = users[i];
    
    
    if(username == ""){ continue; }
    
    var newSheet;
    
    var isExisting = false;
    
    // ### Check if Sheet is already existing ###
    if(!isFirstRun){
      newSheet = searchSheet(username);
      if(newSheet){
        isExisting = true;
        var existingSheetData = newSheet.getDataRange().getValues();
        // * If sheet is existing, update users
//        for (var j = (settings['rowFirstData']-1), k = 0; k < users.length; j++, k++) {    
//          if(existingSheetData.length <= j || existingSheetData[j][letterToColumn(settings['spalteUsernamen'])-1] != users[k]){            
//            newSheet.getRange(j+1, letterToColumn(settings['spalteUsernamen'])).setValue(users[k]);
//          }          
//        }
        // Update complete sheet except grading data
        for (var r = 0; r < rowsN; r++) { 
        	for(var c = 0; c < colsN; c++){
        		//if no grading data -> update if different
        		//Logger.log("Row: "+(r+1)+", Col: "+(c+1));        		
        		//Logger.log("Row length: "+(existingSheetData.length)+", Col length: "+(existingSheetData[0].length));
        		if(!(r > settings['rowFirstData']-2 && c > letterToColumn(settings['colFirstGrade'])-2 && c < colsN-1)){
        			//Logger.log("Nicht im Bereich");
        			//|| existingSheetData[r][c] != vorlageSheetData[r][c]
        			//(existingSheetData.length <= r || existingSheetData[r].length <= c) 
        			if(true){
        				//Logger.log("Vorlage: "+vorlageSheetData[r][c]+", Hier: "+existingSheetData[r][c]);
        				newSheet.getRange(r+1, c+1).setValue(vorlageSheetData[r][c]);
        			}
        		}  
        		//Logger.log("Neue Spalte");
        	}        	
        }
      }        
    }
    
    if(isFirstRun || ! isExisting){
      newSheet = vorlageSheet.copyTo(activeSpreadsheet);    
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
    
  }
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
  }
}

function auswertung() {
  
  if(!isOwner()) return;    
      
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheets=activeSpreadsheet.getSheets();
      
  
  /*var auswertungSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Auswertung");
  if (auswertungSheet == null) {
    Logger.log("Sheet -Auswertung- fehlt");
    return;
  }*/
  
  // Daten von Auswertung durch gehen
  /*var auswertungData = auswertungSheet.getDataRange().getValues();
  for (var i = 1; i < auswertungData.length; i++) {
    Logger.log('User name: ' + auswertungData[i][0]);
  }*/ 
  
  
  // create new Sheet - if exists, delete first  
    activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var yourNewSheet = activeSpreadsheet.getSheetByName(settings['sheetAuswertungName']);

    if (yourNewSheet != null) {
        activeSpreadsheet.deleteSheet(yourNewSheet);
    }

    yourNewSheet = activeSpreadsheet.insertSheet();
    yourNewSheet.setName(settings['sheetAuswertungName']);
    //Logger.log('sheetAuswertungName in Auswertung: ' +settings['sheetAuswertungName']);
  
  activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheets=activeSpreadsheet.getSheets();
  
  var row = 1;
  
  yourNewSheet.getRange(row, 1).setValue("Bewerteter");
  
  row++;
  
  var firstRowOfUserData;
  
  // Walk through sheets to get users  
  for (var i = 0; i < spreadsheets.length; i++) {
     var bewerteter = spreadsheets[i].getName();
     Logger.log('Bewerteter: ' + bewerteter);
     if(bewerteter != settings['sheetAuswertungName'] && bewerteter.indexOf(settings['deletionMarker']) != 0 && bewerteter != settings['nameVorlage'] ){
    	 
       yourNewSheet.getRange(row, 1).setValue(bewerteter).setFontWeight("bold");
       row++;
       
       // ### Headings ###
       // sheetData row in first dimension, col in second
       var sheetData = spreadsheets[i].getDataRange().getValues();
       var sheetName = spreadsheets[i].getName();
       // Create Array for later mean calculation
       var competences = new Array(sheetData[0].length - (letterToColumn(settings['colFirstGrade'])-1));
       
       for (var j = 0; j < competences.length; j++) {
         competences[j]= new Array(sheetData.length);
       }
       Logger.log("competences.length: "+competences.length);
       var weightingHeader = yourNewSheet.getRange(row+1, parseFloat(letterToColumn(settings['colFirstGrade']))-1);
       weightingHeader.setValue(settings['maxPointsText']);
       // Walk through headings
       for (var k = 0; k < sheetData[0].length - (letterToColumn(settings['colFirstGrade'])); k++) {    
    	 //Logger.log("sheetData: "+(settings['rowFirstData']-3)+"; "+(letterToColumn(settings['colFirstGrade'])-1+k));
    	 var heading = sheetData[settings['rowFirstData']-3][letterToColumn(settings['colFirstGrade'])-1+k];
    	 var weighting = sheetData[settings['rowFirstData']-2][letterToColumn(settings['colFirstGrade'])-1+k];
    	 
    	 //Logger.log('Heading: ' + (letterToColumn(settings['colFirstGrade'])+k));
         var headingCell = yourNewSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+k);
         headingCell.setValue("='"+sheetName+"'!"+columnToLetter(letterToColumn(settings['colFirstGrade'])-1+1+k)+(settings['rowFirstData']-2));
         
         var weightingCell = yourNewSheet.getRange(row+1, parseFloat(letterToColumn(settings['colFirstGrade']))+k);
    	 //var headingCell = yourNewSheet.getRange(row, 5);
         //headingCell.setValue("='"+sheetName+"'!"+columnToLetter(letterToColumn(settings['colFirstGrade'])-1+k)+settings['rowFirstData']-3);
         weightingCell.setValue("='"+sheetName+"'!"+columnToLetter(letterToColumn(settings['colFirstGrade'])-1+1+k)+(settings['rowFirstData']-1)+"*"+settings['punkteMax']);
         //headingCell.setFontWeight("bold");
       }
       // Set Text for endrating
       var resultsHeader = yourNewSheet.getRange(row+1, sheetData[0].length);
       resultsHeader.setValue(settings['endRatingText']);
       // Two times for Headers and Weightings
       row++;row++;
       
       firstRowOfUserData = row;
       
       var counterUsers = 0;
       
       // Walk through sheets to get data for user
       for (var j = 0; j < spreadsheets.length; j++) {
         var sheet = spreadsheets[j];
         var bewerter = sheet.getName();
         
         if(bewerter != settings['sheetAuswertungName'] && bewerter != bewerteter && sheet.getName().indexOf(settings['deletionMarker']) != 0 && settings['nameVorlage'] != sheet.getName()){
           
           yourNewSheet.getRange(row, 2).setValue(bewerter);
           
           var sheetData = sheet.getDataRange().getValues();           
           
           // Walk through reviews
           // Walk through lines
           for (var k = 0; k < sheetData.length; k++) {
             // If line is for reviewed person
        	 // Logger.log('Wert1: ' + k);
             //Logger.log('letterToColumn(settings['spalteUsernamen']): '+ sheetData[k][letterToColumn(settings['spalteUsernamen'])-1]+' -- bewerteter: '+bewerteter);
             if(sheetData[k][letterToColumn(settings['spalteUsernamen'])-1] == bewerteter){
               //Logger.log('Bewerteter: ' + bewerteter);
               // calc number of cols with data
               var colsWithData = sheetData[k].length - (letterToColumn(settings['colFirstGrade'])-1);
               // Walk through all cols
               for(var l = 0; l < colsWithData; l++){
                 var val = sheetData[k][parseFloat(letterToColumn(settings['colFirstGrade']))+l-1];
                 
                 // Fill Array for later mean calc
                 // Only if in range of points
                 //Logger.log(l+' -- '+ counterUsers + ' - ' + val+' isInRange: ' );
                 if(isInPointsRange(val) && l < colsWithData -1){
                   competences[l][counterUsers] = val;                 
                 
                   //Logger.log(l+' -- '+ counterUsers + ' - ' + val);
                   // Write Val to cell
                   //var cell = yourNewSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+l).setValue(val);
                   
                 }
                 else{
                   var cell = yourNewSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+l).setValue("");
                 }
                 
                 //Logger.log('Wert2: ' + k);
                 var columnGrade = columnToLetter(parseFloat(letterToColumn(settings['colFirstGrade']))+l);
                 /* set link to user rating value */
                 var cell =  yourNewSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+l).setValue("=IF('"+bewerter+"'!"+columnGrade+(k+1)+" <> \"\"; '"+bewerter+"'!"+columnGrade+(k+1)+" * '"+settings['nameVorlage']+"'!"+columnGrade+(settings['rowFirstData']-1)+";\"\")");
                 
                 /* Calc Row Mean if last col */
                 if(l+1 == colsWithData){               	 
                	                	 
                	 var firstColData = parseFloat(letterToColumn(settings['colFirstGrade']));
                	 var lastColData = parseFloat(letterToColumn(settings['colFirstGrade'])) + colsWithData -2;                	 
                	 var rangeGrades = columnToLetter(firstColData)+row+":"+columnToLetter(lastColData)+row; 
                	 var rangeWeights = columnToLetter(firstColData)+(settings['rowFirstData']-1)+":"+columnToLetter(lastColData)+(settings['rowFirstData']-1);
                	 
                	 var sumWeights = "SUMIFS('"+settings['nameVorlage']+"'!"+rangeWeights+";'"+bewerter+"'!"+columnToLetter(firstColData)+(k+1)+":"+columnToLetter(lastColData)+(k+1)+";\"<>\")";
                	 yourNewSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+l).setValue("=IF("+sumWeights+">0;ROUND(SUM("+rangeGrades+")/"+sumWeights+";2);\"\")");
                	                 	
                 }
                 
                 /* Set Border if not last col */                 
                 else{
                   cell.setBorder(true, true, true, true, true, true);
                 }
               }               
             }  
           }         
         
           row++;
           counterUsers++;
         }
       }
       // Calc mean values
       yourNewSheet.getRange(row, letterToColumn(settings['colFirstGrade'])-1).setValue("Durchschnitt").setFontWeight("bold");
       for (var x = 0; x < competences.length; x++) {
//         var sum = 0;
//         var count = 0;
//         for (var j = 0; j < counterUsers; j++) {
//           var val = competences[x][j];
//           if(isNumeric(val)){
//             sum += val;
//             count ++;
//           }
//           //Logger.log(i+' - '+ j + ' -> ' +  competences[x][j]);
//         }
//         yourNewSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+x).setValue(sum / count).setFontWeight("bold");
    	 var rangeForCalc = columnToLetter(parseFloat(letterToColumn(settings['colFirstGrade']))+x)+firstRowOfUserData+":"+columnToLetter(parseFloat(letterToColumn(settings['colFirstGrade']))+x)+(row-1);
         yourNewSheet.getRange(row, parseFloat(letterToColumn(settings['colFirstGrade']))+x).setValue("=IF(COUNT("+rangeForCalc+")>0;ROUND(AVERAGE("+rangeForCalc+");2);\"\")").setFontWeight("bold");
         
       }
       row++;
       
       
     }
   }
  resizeAllColumns(yourNewSheet, 2);
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