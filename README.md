# Excel-Functions

A collection of Google Sheets functions for reference purposes

Index and Match	
  Example
	=index(U2:U,match(B2,S2:S,0),"")	
  Explainer
	=index(data to return,match(data to look up,matching data that is looked up,0),"")
  
Sumif
Example
  =sumif($P$2:P,A2,$O$2:O)/60
Explainer
  =sumif(Data to reference,Referenced data matched,data added if first two data sets matched)/60
  
  
  Working Days Function
  working days in current month in hours multiplied with daily hours
  Example
  =(NETWORKDAYS(EOMONTH(TODAY(),-1),TODAY()-1))*7.2
  
  
Per Capita Sum
Example
  =Sum(AB2)/countif(MASTERDASHBOARD!$B$2:M,AA2)
Explainer
  =Sum(totalnumber to be devided)/countif(count if this matches, this)	


Queries 
Query to pull data based on Date Range
Example
  =query(DATAPULL!A2:E,"select A, B, C, E where A >= date'2018-07-01'and A <= date'2018-07-31'",0)
Explainer 
	first brackets are which range of a sheet to lookup, Select abce, is which columns to return where column A is 		between the two dates listed.
  =query(DATAPULL!A2:E,"select A, B, C, E where A >= date'2018-07-01'and A <= date'2018-07-31'",0)
  
Query to pull only unique data
Example
=Unique(QUERY({January!C2:C;February!C2:C;March!C2:C;April!C2:C;May!C2:C;June!C2:C;July!C2:C;August!C2:C;September!C2:C;October!C2:C;November!C2:C;December!C2:C},"Select * where Col1 is not null"))
Explainer
unique pulls only unique rows from query
Query selects a column of each sheet defined if Collumn1 is not empty

Importrange 
to import data from another google sheet
Example
=IMPORTRANGE("URLofSheetpasted here","DataConsolidation!A:E")
Explainer 
Copy and paste url of sheet between "" and define the sheet inside "" 

////////// 
=MID(C4,FIND(".ulenscale",C4)-10, 6)

//the bellow query can be used to function as a Vlookup index match/// remember data should be plain text//////
=query($D$1:$E$21, "select E where D ='" & B21 & "' limit 10",0)



////////// used to import specific data based on keyword
=query(IMPORTRANGE("url","ALLUPLOADS!A:I"),"select Col1,Col2,Col3,Col4,Col5,Col6,Col7,Col8,Col9 where Col6='Surf'")



This is the one function used to import, brand data into a brand sheet that should only be viewable by client
=query(IMPORTRANGE("URL of ths Sheet","ALLUPLOADS!A:I"),"select Col1,Col2,Col3,Col4,Col5,Col6,Col7,Col8,Col9 where Col6='Brand Name'")	
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
This extracts the TAB ID, Note that TAB IDs with 6 digits have a space at the end.//////////////////////////////////
=LEFT(B2,7)	
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
This pulls all the User Name Sheets to 'sheet' which is not editable, (so that it won't break). Also it only pulls Unique Rows, so if a user duplicated their Entries then it won't be dupliacted 
=Unique(QUERY({USER.NAME!A2:I;USER.NAME!A2:I},"Select * where Col2 is not null"))

///////////////////////////
=UNIQUE(QUERY({'timeEntries2018-08-29'!A2:T},"Select * where Col1 is not null",0))


//////
need to explain this more but its to insert 8 rows beneath other rows
function insertRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("AXE");/* the name of the sheet*/
  var sourceRange = sheet.getRange(1, 1, sheet.getLastRow());
  var sheetData = sourceRange.getValues();
  for (var i=0; i < sheetData.length; i++) {
    if(sheetData[i][0]!=""){
         sheet.insertRowsAfter(i+1, 8);/*  8 how many rows*/
         sheetData = sheet.getRange(1, 1, sheet.getLastRow()).getValues();
    }
  }
}
////// TERMINAL Merge multiple text files:
cat * > merged-file
found here: https://unix.stackexchange.com/questions/3770/how-to-merge-all-text-files-in-a-directory-into-one
////
finding a word in a string and pulling it in
=QUERY( UniqueAllCombined!A:I, "Select * Where C like '%Group%' ")
Found here
https://productforums.google.com/forum/#!topic/docs/defcUoWf2iI
//
////
remove duplicates in a string (single cell): found:https://stackoverflow.com/questions/50937289/removing-duplicate-strings-from-a-comma-separated-list-in-a-cell
=JOIN(", ",UNIQUE(TRANSPOSE(SPLIT(N2,", "))))




/// A google script that checks if column D is populate and then writes today's date next to it in Column E
/// Original found at https://webapps.stackexchange.com/questions/39086/auto-date-insert-when-opposite-cell-is-populated
//------------------------------------------------------------
 //Auto-Populate date in Column A of when column B is updated 
 //Edited 01/13/16 - MK
  //Auto-Populate date in Column E of when column D is updated 
 //------------------------------------------------------------

function onEdit(event) {
  var eventRange = event.range;
  if (eventRange.getColumn() == 4) { // 2 == column D

    // getRange(row, column, numRows, numColumns)
    var columnXRange = SpreadsheetApp.getActiveSheet().getRange(eventRange.getRow(), 5, eventRange.getNumRows(), 5);

    var values = columnXRange.getValues();

    for (var i = 0; i < values.length; i++) {
      if (!values[i][0]) {  // If cell isn't empty
       values[i][0] = new Date();
      }
    }
    columnXRange.setValues(values);  
  }
}


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
how to find duplicates in google sheets using conditional formatting
found here: https://www.youtube.com/watch?v=skQEKi0zULg
=countif(B:B,B1)>1

