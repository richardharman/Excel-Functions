# Excel-Functions

A collection of Google Sheets functions for reference purposes

Index and Match
	
  Example
	=index(U2:U,match(B2,S2:S,0),"")
	
  Explainer
	=index(data to return,match(data to look up,matching data that is looked up,0),"")
  
Sumif
  =sumif($P$2:P,A2,$O$2:O)/60
  
  working days in current month in hours multiplied with daily hours
  =(NETWORKDAYS(EOMONTH(TODAY(),-1),TODAY()-1))*7.2
  
  
  =Sum(AB2)/countif(MASTERDASHBOARD!$B$2:M,AA2)
  
  Queries used
  =query(DATAPULL!A2:E,"select A, B, C, E where A >= date'2018-07-01'and A <= date'2018-07-31'",0)
  
  
  
  =Unique(QUERY({January!C2:C;February!C2:C;March!C2:C;April!C2:C;May!C2:C;June!C2:C;July!C2:C;August!C2:C;September!C2:C;October!C2:C;November!C2:C;December!C2:C},"Select * where Col1 is not null"))
  
  
  =IMPORTRANGE("URLofSheetpasted here","DataConsolidation!A:E")
