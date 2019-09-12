// File Name:	ZIPEDIT.D
// Purpose:	Used to edit the text symbol of a Zip 
//
// Version:	1.01
// Date:		4.Oct.95
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
//---------------------------------------------------------------------------------------------
//
// 

HANDLE
	hEnt, hTextAsSymbol,
	hDlg,
	hChan;
XY  
	xyInsertPoint, xyNull;
	
STRING 
	sEntityID, sClass, sSymbolName,
	sTmp, sOldText, sNewText ;
NUMBER
	nNull;
	
// Clear current user selection
//
   UserSelection("Clear");
   UserSelection("Update");

// Prompt user to select part of a Zipper
//
   if (! hEnt=GetEntity ("Select Zipper", "all")) Exit (%ok, "User Cancelled") ;

   Display ("cursor", "wait", "Searching for more data");
   GetDBData (hEnt, "ID", &sEntityID);
 
   if (StringLength (sEntityID) <=0) Exit (%ok, "Selected entity is not part of a Zipper!");
   sTmp = "DB ID = '" + sEntityID + "'" ;
   nFound = 0 ;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		sTmp = "" ;
		GetEntityClass(hEnt, &sClass) ;
		GetDBData(hEnt,"Zipper", &sTmp) ;
		if (StringCompare( sTmp, "1") && StringCompare(sClass, "symbol")){
			GetGeometry(hEnt, &sSymbolName, &xyInsertPoint, &xyNull, &nNull) ;
			if (StringCompare(sSymbolName,  "TextAsSymbol")) {
				hTextAsSymbol = hEnt;
				GetDBData (hEnt, "Data", &sOldText);
				}
			}
		}
   	}
   Close("selection", hChan); 
   if (hTextAsSymbol) { 
        nButX = 65; nButY = 35;
        hDlg = Open ("dialog", "Zipper Text Editor", "font Helv 8", 20,20, 215, 60);

	AddControl ( hDlg, "pushbutton",  nButX, nButY, 35, 14, "Cancel", "%cancel", "");
	AddControl ( hDlg, "pushbutton",  nButX + 48 , nButY, 35,14, "OK", "%ok", "");
	AddControl ( hDlg, "ledit", 5, 12, 204, 14 , sOldText , "string", "sNewText");

   	Ok = Display("dialog", hDlg, "%center");
  	Close("dialog", hDlg);
	}

//Change only if Text different
//
   if (!StringCompare(sOldText,sNewText)&&Ok!=%cancel) SetDBData (hTextAsSymbol, "Data", sNewText); 