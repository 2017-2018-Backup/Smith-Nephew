// File Name:	CLUPDATE.D
// Purpose:	To update JOBST CAD system drawings.
//		These drawings will have been created pre-
//		imageABLE 
//
// Version:	1.01
// Date:		08.Jun.95
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
// Notes:-
//	This is an opportunity to clear up some of the 
//	problems caused by developing the full CAD system
//	from the Pilot.
//	

Display ("cursor", "wait", "Updating Drawing");

// Exit if this layer not found as there will be nothing to
// update
  if ( Table("find", "layer", "titlebox") == %badtable) Exit(%ok, "Nothing to update");

HANDLE	
	hDlg,
	hChan,
	hSel,
	hSym,
	hEnt;
XY	
	xySymO,
	xyTitleO,
	xyScale;
	
ANGLE
	aAngle;
	
NUMBER
	nDeltaTitle,
	nDelta;
	
BOOL
	NewArm;
	
STRING 
	sPathJOBST,
	sSymName,	
	sNewSymName,
	sSymSelection;

// Ensure that Layer colour and line style are by layer 
//
   Execute ("menu", "SetStyle",  Table("find", "style", "bylayer")) ;
   Execute ("menu", "SetColor",  Table("find", "color", "bylayer")) ;
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;

// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
	
// Get all of the symbols for updating
//
   sSymSelection = 	"DB SymbolName = 'titlebox' OR " +
   		"DB SymbolName = 'bodybox' OR " +
   		"DB SymbolName = 'sleevebox' OR " +
   		"DB SymbolName = 'armbox' OR " +
   		"DB SymbolName = 'waistbox' OR " +
   		"DB SymbolName = 'legbox' OR " +
   		"DB SymbolName = 'legoptbox' OR " +
   		"DB SymbolName = 'legtitle'" ;
   		
   hChan = Open ("selection", sSymSelection );  
   if (hChan) {
	SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB" );
	nDeltaTitle = - 1.125;
	NewArm = %false;
      	ResetSelection(hChan);
   	while (hSym = GetNextSelection(hChan)) {
   		GetGeometry(hSym, &sSymName, &xySymO, &xyScale, &aAngle);
   		sNewSymName = "";
   		nDelta = nDeltaTitle ;

      		if(StringCompare(sSymName, "titlebox" )){ 
      			sNewSymName = "mainpatientdetails" ;
      			xyTitleO = xySymO;
      			}
   		if(StringCompare(sSymName, "bodybox" )) sNewSymName = "vestbody";
   		if(StringCompare(sSymName, "sleevebox" )) sNewSymName = "vestarm";
  		if(StringCompare(sSymName, "armbox" )){
  			sNewSymName = "armarm";
  			nDelta = nDelta - 0.25;
  			NewArm = %true;
  			}
   		if(StringCompare(sSymName, "waistbox" )) sNewSymName = "waistbody";
   		if(StringCompare(sSymName, "legbox" )) sNewSymName = "waistleg";
   		if(StringCompare(sSymName, "legoptbox" )) sNewSymName = "legleg";
   		if(StringCompare(sSymName, "legtitle" )) sNewSymName = "legcommon";
   		
   		// Revise symbol to new symbol
   		if( !StringCompare(sNewSymName, "")) {
   			Import("symbol",sNewSymName);
   			SetGeometry(hSym, sNewSymName, xySymO.x, xySymO.y + nDelta,  xyScale, aAngle);   				
   			}
   	   	}
   	// If an "armarm" symbol is loaded then we need an "armcommon" symbol
   	//
   	if( NewArm) {
   		Import("symbol","armcommon");
   		hSym = AddEntity("symbol", "armcommon", xyTitleO.x , xyTitleO.y + nDeltaTitle); 
   		SetEntityData( hSym, "layer", Table("find", "layer", "Data")) ;
   		}
	} 
  Close("selection",hChan);

// Delete Existing text on titlebox layer   	
//
 
   hChan = Open ("selection", "layer = 'titlebox' AND type = 'text'");
   if (hChan){
   	UserSelection ("clear");
   	ResetSelection (hChan);
   	nEnt = 0 ;
	while (hEnt = GetNextSelection(hChan))	{
		UserSelection ("add", hEnt);
		nEnt = nEnt + 1 ;
		}
	Close("selection", hChan);
	if (nEnt > 0) Execute ("menu", "Delete"); 
	}
   Close("selection",hChan);
   
// Delete unnessesary DB Fields.
// The limit on the no off DB Fields is about 128 
//
   Table("delete", "field", "Status");
   Table("delete", "field", "ThumbCircum");
   Table("delete", "field", "ThumbLength");
   Table("delete", "field", "NoThumb");
   Table("delete", "field", "DetachGauntlet");
   Table("delete", "field", "EnclosedThumb");
   Table("delete", "field", "FlapStrap");
   Table("delete", "field", "PalmNo");
   Table("delete", "field", "PalmWristDist");
   Table("delete", "field", "SecondTape");
   Table("delete", "field", "SecondLastTape");
   Table("delete", "field", "ShoulderPleat1");
   Table("delete", "field", "ShoulderPleat2");
   Table("delete", "field", "WristNo");
   Table("delete", "field", "WristPleat1");
   Table("delete", "field", "WristPleat2");
   Table("delete", "field", "doctor");
   Table("delete", "field", "measby");
   Table("delete", "field", "orderstatus");
   Table("delete", "field", "pono");
 
   Execute("menu", "ViewRedraw") ;
// End of CLUPDATE.D ;
//
 
