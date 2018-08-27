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
