// File Name:	WOEDIT.D
// Purpose:		Used to edit Work Order Text on a
//			global basis
//
// Version:		1.01
// Date:		11.Oct.95
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
//---------------------------------------------------------------------------------------------
//
// 

HANDLE
	hEnt,
	hDlg,
	hChan;
XY  
	xyInsertPoint, xyScale;
	
STRING 
	sTmp, sOldWO, sNewWO, sOldString, sNewString ;
	
NUMBER
	nn, nAngle, nFound, iOldStringLen , iOldWOLen, iNewWOLen  ;
	
MSGID
	Ok;


function FN_SubstituteText(){
	iOldStringLen = StringLength(sOldString) ;
	if (iOldStringLen == 0 || iOldStringLen < iOldWOLen) return %false;
	nn = 1; 
	while (nn <= (iOldStringLen - iOldWOLen)+1 ) {
		if (StringCompare( sOldWO, StringMiddle(sOldString,nn,iOldWOLen ))){
			sNewString = StringMiddle(sOldString,1, (nn-1)) + sNewWO + StringMiddle(sOldString, nn + iOldWOLen, (iOldStringLen - ( nn + iOldWOLen) + 1 )) ;  
			return %true ;
			}
		nn = nn + 1;			
		}
	return %false;
}

	
// Clear current user selection
//
   UserSelection("Clear");
   UserSelection("Update");
  
// Get old workorder # and new workorder #
   nButX = 35; nButY = 35;
   hDlg = Open ("dialog", "Global Work Order Update", "font Helv 8", 20,20, 150, 60);

   AddControl ( hDlg, "pushbutton",  nButX, nButY, 35, 14, "Cancel", "%cancel", "");
   AddControl ( hDlg, "pushbutton",  nButX + 48 , nButY, 35,14, "OK", "%ok", "");
   AddControl ( hDlg, "static left",   5, 14, 30, 10 , "Old WO:" , "", "");
   AddControl ( hDlg, "ledit",         40, 12, 33, 14 , "" , "string", "sOldWO");
   AddControl ( hDlg, "static left", 75, 14, 32, 10 , "New WO:" , "", "");
   AddControl ( hDlg, "ledit",       110, 12, 33, 14 , "" , "string", "sNewWO");

   Ok = Display("dialog", hDlg, "%center");
   iOldWOLen = StringLength(sOldWO) ;
   iNewWOLen = StringLength(sNewWO) ;

   if (!Ok || iOldWOLen==0 || iNewWOLen == 0 ) {
   	Close("dialog", hDlg);
   	Exit(%ok, "Cancelled or No data");	   
	}
   Close("dialog", hDlg);
   
   Display ("cursor", "wait", "Attempting to global update text with new work order");

   nFound = 0;
   sTmp = "type = 'Text' AND layer = 'Notes'" ;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		sOldString = "" ;
		GetGeometry(hEnt, &sOldString, &xyInsertPoint, &xyScale, &nAngle) ;
		if ( FN_SubstituteText() == %true ) {
			 SetGeometry(hEnt, sNewString, xyInsertPoint, xyScale, nAngle) ;
  			 nFound = nFound + 1;
			 }
		}
      	}
   Close("selection", hChan); 

// Clear current user selection
//
//   UserSelection("Clear");
//   UserSelection("Update");
