heetName = "UK";
  var firstRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange("1:1").getValues()[0];
  Logger.log(firstRow);
  var firstRowDE = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DE").getRange("1:1").getValues()[0];
  
      var userProperties = PropertiesService.getUserProperties();
    //Logger.log(userProperties.getProperty("UK15"));
    //Logger.log(userProperties.getProperty("DE15"));

    userProperties.setProperties({
      "responsibilityCol": String.fromCharCode(65 + firstRow.indexOf("Responsibility")),
        "titleCol": String.fromCharCode(65 + firstRow.indexOf("Title")),
        "descriptionCol": String.fromCharCode(65 + firstRow.indexOf("Description")),
        "bulletsStartCol": String.fromCharCode(65 + firstRow.indexOf("Bullet 1")),
        "bulletsEndCol": String.fromCharCode(65 + firstRow.indexOf("Bullet 5")),
        "searchTermCol": String.fromCharCode(65 + firstRow.indexOf("Search Terms 1")),
      "itemNumberCol": String.fromCharCode(65 + firstRow.indexOf("Item #")),
        "kwTitleCol": String.fromCharCode(65 + firstRow.indexOf("Keywords Title")),
        "kwBaseCol": String.fromCharCode(65 + firstRow.indexOf("Keywords Base")),
        "row-in-progress": "0",
        "dimensionsCol": String.fromCharCode(65 + firstRow.indexOf("Dimensions")),
      "DE.titleCol": String.fromCharCode(65 + firstRowDE.indexOf("Title")),
      "DE.descriptionCol": String.fromCharCode(65 + firstRowDE.indexOf("Description")),
      "DE.bulletsStartCol": String.fromCharCode(65 + firstRowDE.indexOf("Bullet 1")),
      "DE.bulletsEndCol": String.fromCharCode(65 + firstRowDE.indexOf("Bullet 5")),
    });
  
  Logger.log(userProperties.getProperties());
}


function setUpSheet() {

    var initials = "VL";
    var sheetName = "UK";

    var bulletsEndCol = getGlobal("bulletsEndCol");
    var titleCol = getGlobal("titleCol");
    var responsabilityCol = String.fromCharCode(titleCol.charCodeAt(0) - 1);
    var a1range = responsabilityCol + ":" + bulletsEndCol;
    

    var sheetRangeEN = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var sheetDE = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DE").getRange(deRange).getValues();

    var sheetEN = sheetRangeEN.getRange(a1range).getValues();

    var outputContent = [];
    var i, j;
    var emptyRow = ["", "", "", "", "", "", ""];
    var splitter = " || ";

    var userProps = PropertiesService.getUserProperties();
    try {
        i = Number(getGlobal("row-in-progress"));
        Logger.log("Starting with row: " + i);
    } catch (e) {
        Logger.log(e);
        i = 1
    } finally {

        while (i < sheetEN.length) {
            // Logger.log(i);

            //for (i = 1; i < 10; i++) {
            // Logger.log(sheetEN[i][0]);
            //Logger.log(sheetEN[i][0]);
            if (sheetEN[i][0] === initials) {
                try {

                    Logger.log("Cell is for " + initials);
                    sheetDE[i].shift();
                    Logger.log(sheetDE[i]);
                    var rowContent = sheetDE[i].join(splitter);
                    var translated = LanguageApp.translate(rowContent, "de", "en", {
                        contentType: 'html'
                    });
                    var newRow = translated.split(splitter);
                    newRow[1] = htmlFixContent(newRow[1]);
                    outputContent.push(newRow);
                    // return true;
                } catch (e) {
                    Logger.log(e);
                    sheetRangeEN.getRange(titleCol + "0" + ":" + bulletsEndCol + i).setValues(outputContent);
                    userProps.setProperty("row-in-progress", i.toString());
                    // return false;
                } finally {
                    Logger.log("row done");

                }
            } else if (sheetEN[i][0] == "Responsibility") {
                Logger.log("Headings... ignore");
                outputContent.push(sheetEN[i]);
            } else if (sheetEN[i][0].length === 0) {
                outputContent.push(emptyRow);
                break;
            } else {
                outputContent.push(emptyRow);
            }
            userProps.setProperty("row-in-progress", i.toString());

        }
        Logger.log(outputContent);
        var rowInProgress = userProps.getProperty("row-in-progress");
        Logger.log(rowInProgress);
        if (rowInProgress > 0) {
            Logger.log("RANGE > 0. SETTING VALUE OF CELLS");
            var rangeToSet = titleCol + "0" + ":" + bulletsEndCol + rowInProgress;
            Logger.log(rangeToSet);
            //sheetRangeEN.getRange("O2:U10").setValues(outputContent);
            sheetRangeEN.getRange(rangeToSet).setValues(outputContent);
        }


    }



}
function DEtest() {
  var deRange = getGlobal("DE.bulletsStartCol")+":"+ getGlobal("DE.bulletsEndCol");
  Logger.log(deRange);
}

function setUpSheet2() {
    var startTime = (new Date()).getTime();
    var MAX_RUNNING_TIME = 20000; // milliseconds
    var initials = "VL";
    var sheetName = "UK";
    var bulletsEndCol = getGlobal("bulletsEndCol");
    var titleCol = getGlobal("titleCol");
    var responsabilityCol = getGlobal("responsibilityCol");
    var a1range = responsabilityCol + ":" + bulletsEndCol;
    var deRange = getGlobal("DE.titleCol")+":"+ getGlobal("DE.bulletsEndCol");
    var sheetRangeEN = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var sheetDE = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DE").getRange(deRange).getValues();

    var sheetEN = sheetRangeEN.getRange(a1range).getValues();

    var outputContent = [];
    var i, j;
    var emptyRow = ["", "", "", "", "", "", ""];
    var splitter = " || ";
    var userProps = PropertiesService.getUserProperties();
    var startRow = Number(userProps.getProperty('row-in-progress'));
    var MAX_ROW_LOOPS = 25;
    var loopCount = 0;
    // Logger.log(sheetEN[0]);
    for (var ii = startRow; ii < sheetEN.length; ii++) {
        //var currTime = (new Date()).getTime();
        //if (currTime - startTime >= MAX_RUNNING_TIME) {
        if (loopCount < MAX_ROW_LOOPS) {
          Logger.log(sheetEN[ii][0]);

            if (sheetEN[ii][0] === initials) {
                try {

                    Logger.log("Cell is for " + initials);
                    Logger.log(sheetDE[ii]);
                    var rowContent = sheetDE[ii].join(splitter);
                    var translated = LanguageApp.translate(rowContent, "de", "en", {
                        contentType: 'html'
                    });
                    var newRow = translated.split(splitter);
                    newRow[1] = htmlFixContent(newRow[1]);
                    outputContent.push(newRow);
                    // return true;
                } catch (e) {
                    Logger.log(e);
                    // return false;
                } finally {
                    Logger.log("row done");
                  loopCount++;
                }

            } else if (sheetEN[ii][0] == "Responsibility") {
                Logger.log("Headings... ignore");
                sheetEN[ii].shift();
                outputContent.push(sheetEN[ii]);
              
            } else if (sheetEN[ii][0] == 0) {
                outputContent.push(emptyRow);
                break;
            } else {
                outputContent.push(emptyRow);
            }

            userProps.setProperty("row-in-progress", ii.toString());
            
            /*ScriptApp.newTrigger("setUpSheet2")
                     .timeBased()
                     .at(new Date(currTime+REASONABLE_TIME_TO_WAIT))
                     .create();*/

        } else {

            Logger.log("Out of time - saving  . . . ");
            Logger.log(outputContent);
            var rowInProgress = userProps.getProperty("row-in-progress");
            Logger.log(rowInProgress);
            if (rowInProgress > 0) {
                rowInProgress++
                var pasteStartRow;
                if (startRow === 0) {
                    pasteStartRow = startRow + 1;
                } else {
                    pasteStartRow = startRow + 1;
                }
                //startRow++
                Logger.log("RANGE > 0. SETTING VALUE OF CELLS");
                var rangeToSet = titleCol + pasteStartRow + ":" + bulletsEndCol + rowInProgress;
                Logger.log(rangeToSet);
                //sheetRangeEN.getRange("O2:U10").setValues(outputContent);
                sheetRangeEN.getRange(rangeToSet).setValues(outputContent);


            }
            break;
        }

        Logger.log(outputContent);
        var rowInProgress = userProps.getProperty("row-in-progress");
        Logger.log(rowInProgress);
        if (rowInProgress > 0) {
            var pasteStartRow;
            if (startRow === 0) {
                pasteStartRow = startRow + 1;
            } else {
                pasteStartRow = startRow + 1;
            }
            rowInProgress++
            //startRow++
            Logger.log("RANGE > 0. SETTING VALUE OF CELLS");
            var rangeToSet = titleCol + pasteStartRow + ":" + bulletsEndCol + rowInProgress;
            Logger.log(rangeToSet);
            //sheetRangeEN.getRange("O2:U10").setValues(outputContent);
            sheetRangeEN.getRange(rangeToSet).setValues(outputContent);


        }
    }
}

function setGlobalsOld() {

    var userProperties = PropertiesService.getUserProperties();
    //Logger.log(userProperties.getProperty("UK15"));
    //Logger.log(userProperties.getProperty("DE15"));

    userProperties.setProperties({
        "titleCol": "P",
        "descriptionCol": "Q",
        "bulletsStartCol": "R",
        "bulletsEndCol": "V",
        "searchTermCol": "W",
        "kwTitleCol": "X",
        "kwBaseCol": "Y",
        "row-in-progress": "0",
        "dimensionsCol": "I"
    });
}

function getGlobal(propertyName) {
    var userProperties = PropertiesService.getUserProperties();
    return userProperties.getProperty(propertyName);
}

function cacheSheet(sheetName) {
    var userProperties = PropertiesService.getUserProperties();
    var wholeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getDataRange().getValues();
    // var wholeSheetStr = JSON.stringify(wholeSheet);
    var newKV = {};
    for (var i = 0; i < wholeSheet.length; i++) {
        newKV[sheetName + (i + 1)] = JSON.stringify(wholeSheet[i]);
    }

    userProperties.setProperties(newKV);
    // var returnedProp = userProperties.getProperty(sheetName + '49');
    // Logger.log(JSON.parse(returnedProp));
}

function c2n(s) {
    return parseInt(s.charAt(0), 36) - 10;
}

function a1ToPos(a1) {
    var a1 = a1.split(":");
    var res;
    if (a1.length > 1) {
        
        var col1 = c2n(/([A-Z]+?)/gi.exec(a1[0])[0]),
            row1 = Number(/([0-9]+)/gi.exec(a1[0])[0]);
        var col2 = c2n(/([A-Z]+?)/gi.exec(a1[1])[0]),
            row2 = Number(/([0-9]+)/gi.exec(a1[1])[0]);

        var fullRange = {
            "start": [col1, row1],
            "end": [col2, row2]
        };

    } else {
        var col = c2n(/([A-Z]+?)/gi.exec(a1[0])[0]);
        var row = Number(/([0-9]+)/gi.exec(a1[0])[0]);
        var fullRange = {
            "start": [col, row],
            "end": [col, row]
        };
    }
    return fullRange;
}

function getCellFromProps(row, col, sheetName) {
    var uProps = PropertiesService.getUserProperties();
    var cellProp = uProps.getProperty(sheetName + row);
    Logger.log(cellProp);
    return JSON.parse(cellProp)[col];
}

function updateCellFromProps(row, col, newValue, sheetName) {
    sheetName = sheetName || "UK";
    var uProps = PropertiesService.getUserProperties();
    var oldRow = JSON.parse(uProps.getProperty(sheetName + row));
    oldRow.splice(col, 1, newValue);
    uProps.setProperty(sheetName + row, JSON.stringify(oldRow));
}

function forRows(doSomething, cellRange) {

    cellRange = cellRange || SpreadsheetApp.getActiveRange();
    var sheet = SpreadsheetApp.getActive();

    var notation = cellRange.getA1Notation();
    var fullRange = a1ToPos(notation);
    var newCellRange = "A" + fullRange['start'][1].toString() + ":" + "AA" + fullRange['end'][1].toString() ;
    
    var contentToFix = sheet.getRange(newCellRange).getDisplayValues();

    var outputContent = [];
    
    var i;

    for (i = 0; i < cellRange.getHeight(); i++) {

        var newRow = [];
        //Logger.log(contentToFix[i]);
        newRow = doSomething(contentToFix[i]);

        outputContent.push(newRow);
    }
    return outputContent;

}



function getDimensions(){
  var sheet = SpreadsheetApp.getActive();
  var rangeToFix = SpreadsheetApp.getActiveRange();
    var dimensionsCol = getGlobal("dimensionsCol");
    var descriptionCol = getGlobal("descriptionCol");

    var reg = /<li>(.*?)<\/li>/g;
  
    // var descriptionData = sheet.getRange(descriptionCol + rowNumber).getDisplayValue();
  var dimensions = "";
   var matches = [];
   var match;

  var dimensionsIntro = "Dimensions:";
  var dimensionsList = forRows(function(row){
    
    //Logger.log(row);
    //Logger.log("DimensionIndex: " + descriptionIndex);
    //Logger.log("DATA: " + row[descriptionIndex]);
    var descriptionIndex = c2n(descriptionCol);
    var descriptionData = row[descriptionIndex].toString();
    while (match = reg.exec(descriptionData)) {
      
               match = match[1]
               if (match.indexOf(dimensionsIntro) !== -1) {
                 dimensions = match.split(dimensionsIntro)[1].trim();
          
                 return [dimensions]
                 break;
               }
                
    }
    
    
  
  }, rangeToFix);
  //sheet.getA(dimensionsList);

   //sheet.getRange(dimensionsCol + rowNumber).setValue(dimensions);

}

function removeKeywordConflicts() {
    
    var sheet = SpreadsheetApp.getActive();
    var kwTitleIndex = c2n(getGlobal("kwTitleCol"));
    var kwBaseIndex = c2n(getGlobal("kwBaseCol"));
    
    Logger.log(kwTitleIndex);
    //Logger.log(kwBaseIndex);
    
    var newRows = forRows(function(row){
      var keywordsBase = row[kwBaseIndex];
      var keywordsTitle = row[kwTitleIndex];
      
      keywordsTitle.map(function(word){
        var indexKw = keywordsBase.indexOf(word);
        if (indexKw > -1) {
          keywordsBase.splice(indexKw, 1);
        }
        
      });
      
      row[kwBaseIndex] = keywordsBase;  
      return row;
    });
  
  
  
   
  
}

function forCells(doSomething, cellRange) {

    cellRange = cellRange || SpreadsheetApp.getActiveRange();

    var contentToFix = cellRange.getDisplayValues();

    var outputContent = [];
    var i, j;

    for (i = 0; i < cellRange.getHeight(); i++) {

        var newRow = [];

        for (j = 0; j < cellRange.getWidth(); j++) {

            var content = contentToFix[i][j]

            newRow[j] = doSomething(content);

        }
        outputContent.push(newRow);
    }
    return outputContent;
}

function regexEscape(str) {
    return str.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
}

function findKeywords() {

    var sheet = SpreadsheetApp.getActive();
    var rowNumber = SpreadsheetApp.getActiveRange().getRow();

    var contentNot = getGlobal("titleCol") + rowNumber + ":" + getGlobal("bulletsEndCol") + rowNumber;
    var contentRange = sheet.getRange(contentNot);
    var kwBaseCell = getGlobal("kwBaseCol") + rowNumber;
    var kwTitleCell = getGlobal("kwTitleCol") + rowNumber;

    var unmarkedContent = sheet.getRange(contentNot).getDisplayValues();
    var kwBase = sheet.getRange(kwBaseCell).getDisplayValue().split(", ");
    var kwTitle = sheet.getRange(kwTitleCell).getDisplayValue().split(", ");

    var kwds = kwBase.concat(kwTitle);

    kwds.sort(function (a, b) {
        return b.length - a.length;
    });

    Logger.log("\n\n" + kwds + "\n");

    var markedContent = forCells(function (content) {
        var reg, regStr, newContent;
        kwds.map(function (kw, i) {
            regStr = "([^$])(" + regexEscape(kw) + ")([^$])";
            reg = RegExp(regStr, "ig");

            /* Logger.log(kw);
            Logger.log(content.replace(kw, "$" + kw + "$"));*/


            content = content.replace(reg, function (a, b, c, d) {
                // Logger.log(" $" + b + "$ ");
                return b + "$" + c + "$" + d;
            });

        });

        return content;

    }, contentRange);

    Logger.log(markedContent);
    contentRange.setValues(markedContent);

}

function retrieveDimensions() {
  var sheet = SpreadsheetApp.getActive();
  var rowNumber = SpreadsheetApp.getActiveRange().getRow();
    var dimensionsCol = getGlobal("dimensionsCol");
    var descriptionCol = getGlobal("descriptionCol");

    var reg = /<li>(.*?)<\/li>/g;
  
    var descriptionData = sheet.getRange(descriptionCol + rowNumber).getDisplayValue();
  var dimensions = "";
   var matches = [];
   var match;
  Logger.log("DATA:   "+descriptionData);
  var dimensionsIntro = "Dimensions:";
    while (match = reg.exec(descriptionData)) {
      
               match = match[1]
               if (match.indexOf(dimensionsIntro) !== -1) {
                 dimensions = match.split(dimensionsIntro)[1].trim();
                 break;
               }
                
            }

   sheet.getRange(dimensionsCol + rowNumber).setValue(dimensions);


}

function getAllDimensions() {
  var sheet = SpreadsheetApp.getActive();
  var descriptionCol = getGlobal("descriptionCol");
  var dimensionsCol = getGlobal("dimensionsCol");
  var descriptionRange = sheet.getRange(descriptionCol+":"+descriptionCol);
  var descriptionVals = sheet.getRange(descriptionCol+":"+descriptionCol).getValues();
  //Logger.log(descriptionVals);
  var dimensionsOutput = [];
  var i;
  for (i=1; i<descriptionVals.length; i++) {
    
    if (descriptionVals[i][0].length > 0) {
      // get dimensions
      var reg = /<li>(.*?)<\/li>/g;
      var dimensionsIntro = "Dimensions:";
      var measurementsIntro = "Measurements:";
      var sizeIntro = "Size:";
      var matches = [];
      var match;
      
      while (match = reg.exec(descriptionVals[i])) {
      
        match = match[1]
        if (match.indexOf(dimensionsIntro) !== -1) {
          var dimensions = match.split(dimensionsIntro)[1].trim();
          //Logger.log(dimensions);
        } else if (match.indexOf(measurementsIntro) !== -1){
          var dimensions = match.split(measurementsIntro)[1].trim();
        } else if (match.indexOf(sizeIntro) !== -1){
          var dimensions = match.split(sizeIntro)[1].trim();
        }
                
      }
      var dimensions = dimensions || "";
      Logger.log(dimensions);
      dimensionsOutput.push([dimensions]);
      
      // append to 3d array
      
    } else {
      Logger.log("BREAKING @ "+ i);
      break;
    }
    
  }
  Logger.log(dimensionsOutput);
  sheet.getRange(dimensionsCol+"2"+":"+dimensionsCol+(dimensionsOutput.length+1)).setValues(dimensionsOutput);

}

function correctItemNumbers(){
  var sheet = SpreadsheetApp.getActive();
  var descriptionCol = getGlobal("descriptionCol");
  var itemNumberCol = getGlobal("itemNumberCol");
  //Logger.log(descriptionCol+"1:"+descriptionCol);
  var descriptionRange = sheet.getRange(descriptionCol+"1:"+descriptionCol);
  var descriptionVals = sheet.getRange(descriptionCol+":"+descriptionCol).getValues();
  var itemNumberVals = sheet.getRange(itemNumberCol+":"+itemNumberCol).getValues();
  var flatItemNumberArray = [].concat.apply([], itemNumberVals)
  var reg = /(<li>MODEL NUMBER:)(.*?)(<\/li>)/gi;
  //Logger.log(flatItemNumberArray);
  var i=0;
  var newDescriptions = forCells(function(content){
    //Logger.log("    FIRST    "+itemNumberVals);
    //Logger.log(itemNumberVals[i][0]);
    if (flatItemNumberArray[i] === undefined || content === "Description" || content.length < 1){
      i++;
      return content;
    } else {
      if (content.toUpperCase().indexOf("MODEL NUMBER:") !== -1){
        var replaced = content.replace(reg, function(a,b,c,d){
          return b + " " + flatItemNumberArray[i] + " " + d;
        });
        i++;
        return replaced;
      } else {
        var modelNumberListItem = "<li>Model Number: " + flatItemNumberArray[i] + " </li>";
        var appended = content.slice(0, content.indexOf("</ul>")) + modelNumberListItem + "</ul>";
        i++;
        return appended;
      }
    }
    
  },
  descriptionRange); 
  Logger.log(newDescriptions);
  descriptionRange.setValues(newDescriptions);
  
}

function retrieveAllKeywords() {

    var sheet = SpreadsheetApp.getActive();
    var descriptionCol = getGlobal("descriptionCol");
    var bulletsEndCol = getGlobal("bulletsEndCol");
    var titleCol = getGlobal("titleCol");
    var kwBaseCol = getGlobal("kwBaseCol");
    var kwTitleCol = getGlobal("kwTitleCol");


    var rowNumber = 110;
    while (rowNumber < 152) {


        var kwBaseRange = descriptionCol + rowNumber + ":" + bulletsEndCol + rowNumber;
        var kwTitleRange = titleCol + rowNumber;

        var kwBaseCell = kwBaseCol + rowNumber;
        var kwTitleCell = kwTitleCol + rowNumber;

        var ranges = {};
        ranges[kwBaseRange] = kwBaseCell;
        ranges[kwTitleRange] = kwTitleCell;

        var allMatches = [];

        for (r in ranges) {
            Logger.log(r);
            var textToSearch = sheet.getRange(r).getDisplayValues();
            var matches = [];
            var re = /\$(.*?)\$/g;
            var match;

            while (match = re.exec(textToSearch)) {
                matches.push(match[1]);
            }

            sheet.getRange(ranges[r]).setValue(matches.join(", "));

            allMatches.push(matches)
        }

        rowNumber++;

    }


}

function retrieveKeywords() {

    var sheet = SpreadsheetApp.getActive();

    var rowNumber = SpreadsheetApp.getActiveRange().getRow();

    var descriptionCol = getGlobal("descriptionCol");
    var bulletsEndCol = getGlobal("bulletsEndCol");
    var titleCol = getGlobal("titleCol");
    var kwBaseCol = getGlobal("kwBaseCol");
    var kwTitleCol = getGlobal("kwTitleCol");

    var kwBaseRange = descriptionCol + rowNumber + ":" + bulletsEndCol + rowNumber;
    var kwTitleRange = titleCol + rowNumber;

    var kwBaseCell = kwBaseCol + rowNumber;
    var kwTitleCell = kwTitleCol + rowNumber;

    var ranges = {};
    ranges[kwBaseRange] = kwBaseCell;
    ranges[kwTitleRange] = kwTitleCell;

    var allMatches = [];

    for (r in ranges) {
        Logger.log(r);
        var textToSearch = sheet.getRange(r).getDisplayValues();
        var matches = [];
        var re = /\$(.*?)\$/g;
        var match;

        while (match = re.exec(textToSearch)) {
            matches.push(match[1]);
        }

        sheet.getRange(ranges[r]).setValue(matches.join(", "));

        allMatches.push(matches);
    }
    return allMatches;
}

function removeDups() {

    var sheet = SpreadsheetApp.getActive();
    var dupRange = SpreadsheetApp.getActiveRange();

    var updatedContent = forCells(function (content) {

        a = content.split("; ");

        uniqueArray = a.filter(function (item, pos, self) {
            return self.indexOf(item) == pos;
        });

        return uniqueArray.join("; ");

    }, dupRange);

    dupRange.setValues(updatedContent);

}

function roundDecs() {
    var rangeToFix = SpreadsheetApp.getActiveRange();
    var re = /([0-9]*\.[0-9]{2})|[^\.]([0-9]+?) cm/g;
    var roundedContent = forCells(function (content) {
        var replaced = content.replace(re, function (a, b, c) {
          
          if (b !== undefined){
            splitDecimal = b.split(".");
            var decimalPlaces = Math.round(Number(splitDecimal[1]) / 10) * 10;
            var wholeNum = Number(splitDecimal[0]);
            var decimalPlacesString = decimalPlaces.toString();
            if (decimalPlaces === 100) {
                return (wholeNum + 1).toString() + ".0";
            } else if (decimalPlaces < 5) {
                return wholeNum.toString() + ".0";
            } else {
                return wholeNum + "." + decimalPlacesString[0]
            } 
          } else {
             return " " + c + ".0 cm";
          }
          
        });
        return replaced;
    });
    rangeToFix.setValues(roundedContent);
}

function stripSpace() {
  
  var rangeToFix = SpreadsheetApp.getActiveRange();
  var stripped = forCells(function (content){
    
    if (content.charAt(0) === '"' && content.charAt(content.length -1) === '"') {
      return content.substr(1,content.length -2).trim();
    } else {
      return content.trim()
    }
    
    
  }, rangeToFix);
  
  rangeToFix.setValues(stripped);
  
}

function htmlFixContent(content) {

    var x = content.lastIndexOf('</li>');
    if (x != -1) {
        content = content.substr(0, x) + content.substr(x) + "\n";
    }
    var y = content.indexOf('<li>');
    if (y != -1) {
        content = content.substr(0, y) + "\n\n" + content.substr(y);
    }
    var z = content.indexOf('<ul>');
  if (z === -1) {
    content = '<ul>' + content.substr((content.indexOf('<li>')-1), -1)
  }

    // "$1"+"\n"+"$2"
    var replaced = content.replace(/(<\/\w?\w?>)(<\w)/g, function (a, b, c) {
        if (b === '</li>') {
            return b + "\n" + c
        } else {
            return b + "\n\n" + c
        }
    }).replace("Color", "Colour").replace(/(<li>\s)([a-z])/g, function (a, b, c) {
        return b + c.toUpperCase();
    }).replace(/(>) /g, ">");
    return replaced;

}

function parseFix() {
  var parser = new DOMParser();
  parser.parseFromString(content, "text/html");
}
// Works on title
function capitalsFixTitle() {
  // Replaces any lower case letters in title
  var rangeToFix = SpreadsheetApp.getActiveRange();
  //var reg = /(<\w*>)([a-z])/g;
  
  

    var capitalisedContent = forCells(function (content) {
      
      
        //return content.replace(reg, function(a,b,c){ return b + c.toUpperCase(); });
      
      
      var words = content.split(" ");
      var exclusionWords = [ "of", "for", "a", "as", "such", "with", "and", "or", "cm", "the", "to", "in", "mdesign", "x"];
      var capitalised = words.map(function (wordObject, index) {

        var word = wordObject.toString();
        
        var firstLetterCode = word.charCodeAt(0);
                        if (!exclusionWords.includes(word.toLowerCase())) {

                        if (firstLetterCode > 96 && firstLetterCode < 123) {
          
                          //Logger.log(word.charAt(0).toUpperCase() + word.slice(1));
                          return word.charAt(0).toUpperCase() + word.slice(1);
        
                        } else {
                           return word;
                        }
                        
                        } else if (words[index-1] === "-"){
                          
                          return word.charAt(0).toUpperCase() + word.slice(1);
                        
                        } else {
                          Logger.log(word);
                          Logger.log(word.charAt(0).toLowerCase() + word.slice(1));
                          return word.charAt(0).toLowerCase() + word.slice(1);
                        }
        
      });
      return capitalised.join(" ");
    }, rangeToFix);
  
  rangeToFix.setValues(capitalisedContent);

}

// Works on description
function capitalsFix() {
  // Replaces any lower case letters after any html tag
  var rangeToFix = SpreadsheetApp.getActiveRange();
  var reg = /(<\w*>)([a-z])/g;

    var capitalisedContent = forCells(function (content) {
        return content.replace(reg, function(a,b,c){ return b + c.toUpperCase(); });
    }, rangeToFix);
  
  rangeToFix.setValues(capitalisedContent);

}

function htmlFix(rangeToFix) {

    rangeToFix = rangeToFix || SpreadsheetApp.getActiveRange();

    var fixedContent = forCells(function (content) {

        var x = content.lastIndexOf('</li>');
        if (x != -1) {
            content = content.substr(0, x) + content.substr(x) + "\n";
        }
        var y = content.indexOf('<li>');
        if (y != -1) {
            content = content.substr(0, y) + "\n\n" + content.substr(y);
        }
        var z = content.indexOf('<ul>');
  if (z === -1) {
    content = content.substr(0, y) + '\n<ul>\n' + content.substr(y);
  }
      var q = content.lastIndexOf('</ul>');
      if (q === -1) {
        content = content.substr(0, x) + content.substr(x) + "</ul>\n";
      }

        // "$1"+"\n"+"$2"
        var replaced = content.replace(/(<\/\w?\w?>)(<\w)/g, function (a, b, c) {
            if (b === '</li>') {
                return b + "\n" + c
            } else {
                return b + "\n\n" + c
            }
        }).replace("Color", "Colour").replace(/(<li>\s)([a-z])/g, function (a, b, c) {
            return b + c.toUpperCase();
        }).replace(/(>) /g, ">");
        return replaced;

    }, rangeToFix);

    rangeToFix.setValues(fixedContent);

}

function splitByLine() {

    var rangeToFix = SpreadsheetApp.getActiveRange();

    var fixedContent = forCells(function (content) {
      try{
        return content.match(/[^\r\n]+/g).join("; ");
      } catch (e){
        return content;
        }
    }, rangeToFix);

    rangeToFix.setValues(fixedContent);

}

function splitByLineComma() {

    var rangeToFix = SpreadsheetApp.getActiveRange();
    var fixedContent = forCells(function (content) {
        return content.split("\n").join(", ");
    }, rangeToFix);

    rangeToFix.setValues(fixedContent);

}

function expandSemiColon() {

    var rangeToFix = SpreadsheetApp.getActiveRange();

    var fixedContent = forCells(function (content) {
        return content.split("; ").join("\n");
    }, rangeToFix);

    rangeToFix.setValues(fixedContent);

}

function searchTermGenerate() {
    var sheet = SpreadsheetApp.getActive();

    var rowNumber = SpreadsheetApp.getActiveRange().getRow();
    var searchTermCol = getGlobal("searchTermCol");
    var kwBaseCol = getGlobal("kwBaseCol");
    var kwTitleCol = getGlobal("kwTitleCol");
    var kwMulti = sheet.getRange(kwTitleCol + rowNumber + ":" + kwBaseCol + rowNumber).getDisplayValues()[0];
    var contentKeywords = [].concat.apply([], kwMulti);
  var optimisedSearchTerms = sheet.getRange(kwTitleCol + rowNumber).getDisplayValues()[0].toString().split(", ");
    var searchTerms = sheet.getRange(searchTermCol + rowNumber).getDisplayValue().match(/[^\r\n]+/g);
    Logger.log("Fixing search terms for row: " + rowNumber + "\n\n");
  

    Logger.log(optimisedSearchTerms);

    var uniqueKeywords = searchTerms.filter(function (item, pos, self) {
        if (!contentKeywords.includes(item)) {
            return self.indexOf(item) == pos;
        }
    });
    
    uniqueKeywords.map(function (keyword) {

        var subKeywords = keyword.split(" ");

        subKeywords.map(function (subKeyword) {

            if (!optimisedSearchTerms.includes(subKeyword)) {
                Logger.log(subKeyword);
                var searchTermLength = optimisedSearchTerms.join(";").replace(/\s/g, '').length;
                if (searchTermLength < 230) {
                    optimisedSearchTerms.push(subKeyword);
                }
            }

        });

    });
    Logger.log(optimisedSearchTerms);
    while (true) {
        var searchTermLength = optimisedSearchTerms.join(";").replace(/\s/g, '').length;
        if (searchTermLength < 230) {
            var randomIndex = Math.floor(Math.random() * (uniqueKeywords.length));
            if (!optimisedSearchTerms.includes(uniqueKeywords[randomIndex])) {
                optimisedSearchTerms.push(uniqueKeywords[randomIndex]);
            }
            Logger.log(uniqueKeywords[randomIndex]);
        } else {
            break;
        }
    }

    sheet.getRange(searchTermCol + rowNumber).setValue(optimisedSearchTerms.join("; "));
}

function searchTermsFix() {

    var sheet = SpreadsheetApp.getActive();

    var rowNumber = SpreadsheetApp.getActiveRange().getRow();
    var searchTermCol = getGlobal("searchTermCol");
    var searchTerms = sheet.getRange(searchTermCol + rowNumber).getDisplayValue().match(/[^\r\n]+/g);
    Logger.log("Fixing search terms for row: " + rowNumber + "\n\n");

    if (searchTerms.length > 1) {

        Logger.log("Complete array of search terms\n ---------- \n " + searchTerms + "\n ---------- \n");

        /*searchTerms.sort(function (a, b) {
            // ASC  -> a.length - b.length
            // DESC -> b.length - a.length
            return b.length - a.length;
          }); */
        var uniqueKeywords = searchTerms.filter(function (item, pos, self) {
            if (!retrieveKeywords().includes(item)) {
                return self.indexOf(item) == pos;
            }
        });
        var optimisedSearchTerms = [];
        uniqueKeywords.map(function (keyword) {

            var subKeywords = keyword.split(" ");

            subKeywords.map(function (subKeyword) {

                if (!optimisedSearchTerms.includes(subKeyword)) {
                    Logger.log(subKeyword);
                    var searchTermLength = optimisedSearchTerms.join(";").replace(/\s/g, '').length;
                    if (searchTermLength < 230) {
                        optimisedSearchTerms.push(subKeyword);
                    }
                }

            });

        });
        Logger.log(optimisedSearchTerms);
        while (true) {
            var searchTermLength = optimisedSearchTerms.join(";").replace(/\s/g, '').length;
            if (searchTermLength < 230) {
                var randomIndex = Math.floor(Math.random() * (uniqueKeywords.length));
                if (!optimisedSearchTerms.includes(uniqueKeywords[randomIndex])) {
                    optimisedSearchTerms.push(uniqueKeywords[randomIndex]);
                }
                Logger.log(uniqueKeywords[randomIndex]);
            } else {
                break;
            }
        }

        sheet.getRange(searchTermCol + rowNumber).setValue(optimisedSearchTerms.join("; "));

    } else {
        Logger.log("No search terms found in cell: " + searchTermCol + rowNumber);
        return;
    }

}

function translateCells(cb) {
    cb = cb || true
    var rangeToTranslate = SpreadsheetApp.getActiveRange();

    var translatedContent = forCells(function (content) {
        return LanguageApp.translate(content, "de", "en", {
            contentType: 'html'
        });
    }, rangeToTranslate);

    rangeToTranslate.setValues(translatedContent);
    return cb
}

function getChars() {
    SpreadsheetApp.getActiveSpreadsheet().toast("" + SpreadsheetApp.getActiveRange().getValues()[0][0].length + " Characters", "Length", 5);
}

function showTitlesLessThan180() {
  
  var titleCol = getGlobal('titleCol');
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
  var titleRange = sheet.getRange(titleCol  + "1:" + titleCol +"60");
  var i = 1;
  
  forCells(function(content){
    if (i>2) {
      if (content.length > 2 && content.length < 180) {
        sheet.getRange(titleCol + i).setBackground('red');
      } else if (content.length > 200){
        sheet.getRange(titleCol + i).setBackground('blue');
      } else {
        sheet.getRange(titleCol + i).setBackground('white');
      }
    }
    
    i++;
  
  },
  titleRange);
  
}

function removeDollars() {
    var activeRange = SpreadsheetApp.getActiveRange();
    activeRange.setValues(forCells(function (content) {
        return content.replace(/\$/g, "");
    }, activeRange));
}

function getTranslations(cb) {
    cb = cb || true
    var sheet = SpreadsheetApp.getActive();
    var rowNumber = SpreadsheetApp.getActiveRange().getRow();

    var bulletsEndCol = getGlobal("bulletsEndCol");
    var titleCol = getGlobal("titleCol");
    Logger.log(bulletsEndCol);
    Logger.log(titleCol);


    var cellRange = titleCol + rowNumber + ":" + bulletsEndCol + rowNumber;
    //var sheetToLeft = (sheet.getSheets().length - 2);
    // Logger.log(sheetToLeft);
    //var translationSheet = sheet.getSheets()[sheetToLeft];
    var translationSheet = sheet.getSheetByName("DE");
    rangeToCopy = translationSheet.getRange(cellRange);
    rangeData = rangeToCopy.getDisplayValues();
    var pasteRange = titleCol + rowNumber + ":" + bulletsEndCol + rowNumber;
    sheet.getRange(pasteRange).setValues(rangeData);
    return cb;
}

function setUpSheet() {

    var initials = "VL";
    var sheetName = "UK";

    var bulletsEndCol = getGlobal("bulletsEndCol");
    var titleCol = getGlobal("titleCol");
    var responsabilityCol = String.fromCharCode(titleCol.charCodeAt(0) - 1);
    var a1range = responsabilityCol + ":" + bulletsEndCol;

    var sheetRangeEN = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var sheetDE = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DE").getRange(a1range).getValues();

    var sheetEN = sheetRangeEN.getRange(a1range).getValues();

    var outputContent = [];
    var i, j;
    var emptyRow = ["", "", "", "", "", "", ""];
    var splitter = " || ";

    var userProps = PropertiesService.getUserProperties();
    try {
        i = Number(getGlobal("row-in-progress"));
        Logger.log("Starting with row: " + i);
    } catch (e) {
        Logger.log(e);
        i = 0
    } finally {

        for (i = i; i < sheetEN.length; i++) {
            // Logger.log(i);

            //for (i = 1; i < 10; i++) {
            // Logger.log(sheetEN[i][0]);
            //Logger.log(sheetEN[i][0]);
            if (sheetEN[i][0] === initials) {
                try {

                    Logger.log("Cell is for " + initials);
                    sheetDE[i].shift();
                    Logger.log(sheetDE[i]);
                    var rowContent = sheetDE[i].join(splitter);
                    var translated = LanguageApp.translate(rowContent, "de", "en", {
                        contentType: 'html'
                    });
                    var newRow = translated.split(splitter);
                    newRow[1] = htmlFixContent(newRow[1]);
                    outputContent.push(newRow);
                    // return true;
                } catch (e) {
                    Logger.log(e);
                    sheetRangeEN.getRange(titleCol + "0" + ":" + bulletsEndCol + i).setValues(outputContent);
                    userProps.setProperty("row-in-progress", i.toString());
                    // return false;
                } finally {
                    Logger.log("row done");

                }

            } else if (sheetEN[i][0] == 0) {
                outputContent.push(emptyRow);
                break;
            } else {
                outputContent.push(emptyRow);
            }
            userProps.setProperty("row-in-progress", i.toString());

        }
        Logger.log(outputContent);
        var rowInProgress = userProps.getProperty("row-in-progress");
        Logger.log(rowInProgress);
        if (rowInProgress > 0) {
            Logger.log("RANGE > 0. SETTING VALUE OF CELLS");
            var rangeToSet = titleCol + 1 + ":" + bulletsEndCol + rowInProgress;
            Logger.log(rangeToSet);
            //sheetRangeEN.getRange("O2:U10").setValues(outputContent);
            sheetRangeEN.getRange(rangeToSet).setValues(outputContent);
        }

    }



}

function onEdit(e) {
    var r = e.range.getRow();
    var c = e.range.getColumn();
    var val = e.value;

    var bulletsRange = ["Q", "R", "S", "T", "U"];

    if (bulletsRange.includes(c) && val.length > 250) {

        e.range.setBackgroundColor("#f47a42");

    }
    SpreadsheetApp.getActiveSpreadsheet().toast("" + val.length + " Characters", "Length", 5);
    updateCellFromProps(r, c, val);
    // var storedVal = getCellFromProps(r, c, "UK");
    // Logger.log(storedVal);
    // SpreadsheetApp.getActiveSpreadsheet().toast(storedVal);
}

function onOpen() {
    var spreadsheet = SpreadsheetApp.getActive();
    // cacheSheet("DE");
    var menuItems = [
        {
            name: 'SET UP SHEET',
            functionName: 'setUpSheet2'
        },
        {
            name: 'Fetch keywords from this row',
            functionName: 'retrieveKeywords'
        },
        {
            name: 'Show keywords in his row with $$',
            functionName: 'findKeywords'
        },
        {
            name: 'Fetch translations from other sheet for this row',
            functionName: 'getTranslations'
        },
      {
        name: "Fix capitals",
        functionName: 'capitalsFix'
        
      },
        {
            name: 'Display length in characters of current cell',
            functionName: 'getChars'
        },
        {
            name: 'Split by line -> ;',
            functionName: 'splitByLine'
        },
        {
            name: 'Split by line -> ,',
            functionName: 'splitByLineComma'
        },
              {
            name: 'Fix caps in Title',
            functionName: 'capitalsFixTitle'
        },
      {
        name: "Get dimensions for this row",
        functionName: "retrieveDimensions"
      },
      {
        name: "Get all dimensions",
        functionName: "getAllDimensions"
      },
        {
            name: 'Remove duplicate items separated by ;',
            functionName: 'removeDups'
        },
        {
            name: 'Translate currently selected cells',
            functionName: 'translateCells'
        },
        {
            name: 'Fix search terms if keywords already done',
            functionName: 'searchTermGenerate'
      },
      {
        name: "Strip white space & quotes from end of cells",
        functionName: "stripSpace" 
        
      },
        {
            name: 'Fix search terms for this row',
            functionName: 'searchTermsFix'
        },
        {
            name: 'HTML->Readable',
            functionName: 'htmlFix'
        },
      {
        name: "Correct all item #",
        functionName: "correctItemNumbers"
      },
        {
            name: 'Retrieve all keywords from all rows',
            functionName: 'retrieveAllKeywords'
        },
        {
            name: 'Remove dollar signs on keywords',
            functionName: 'removeDollars'
        },
        {
            name: "Round currently selected cells",
            functionName: "roundDecs"
      },
        {
            name: "Expand semi colon with \n",
            functionName: "expandSemiColon"
      }
  ];
    spreadsheet.addMenu('Remazing Helper', menuItems);

    setGlobals();
    /* ScriptApp.newTrigger("onChange")
      .forSpreadsheet(sheet)
      .onEdit()
      .create();*/
}

if (![].includes) {
    Array.prototype.includes = function (searchElement /*, fromIndex*/ ) {
        'use strict';
        var O = Object(this);
        var len = parseInt(O.length) || 0;
        if (len === 0) {
            return false;
        }
        var n = parseInt(arguments[1]) || 0;
        var k;
        if (n >= 0) {
            k = n;
        } else {
            k = len + n;
            if (k < 0) {
                k = 0;
            }
        }
        var currentElement;
        while (k < len) {
            currentElement = O[k];
            if (searchElement === currentElement ||
                (searchElement !== searchElement && currentElement !== currentElement)) {
                return true;
            }
            k++;
        }
        return false;
    };
}







