// File Name:	CL_UPDTE.D
// Purpose:	To update JOBST CAD system drawings.
//			These drawings will have been created pre imageABLE 
//
// Version:	1.01
// Date:		30.Oct.95
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
// 09.Dec.96	GG	Warn of inconsistent units when updating a Drawing
//				
// Notes:-
//	This is an opportunity to clear up some of the 
//	problems caused by developing the full CAD system
//	from the Pilot.
//	
// The following information will be provided from the calling macro:-
//
//	sFileNo
//	sAge
//	sDiagnosis	
//	sPatient
//	sWorkOrder
//	sOrderDate
//	sSex
//	sUnits
//	sTemplateEngineer
//	sCADFile
// N.B.
// 	This Macro is called from a Macro Created by the VISUAL BASIC
//	programme CLTXFDIA.EXE
//	The patient data above will have been read from the TXF file
//	associated with this order.
//	The existing CAD file will have been selected by the user.
//
//

   Display ("cursor", "wait", "Updating Drawing");

HANDLE	
	hDlg,
	hChan,
	hSel,
	hSym,
	hMPD,
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
	sOldFileNo,
	sOldUnits,
	sText,
	sSide,
	sFabric,
	sDocFile,
	sDBData,
	sNewDBData,
	sPathJOBST,
	sSymName,	
	sNewSymName,
	sEntSelection,
	sSymSelection;

   
// Get Existing drawing name then close it
//
   GetData ("DocFile", &sDocFile);
   SetData ("FileChanged", 0);	
   Execute ("menu", "FileClose");

// Get existing drawing and update it  
//
  if (!Import ("drawing", sCADFile)) Exit (%abort, "Can't import Old DRAWING!");  
    
// Exit if this layer not found as there will be nothing to
// update
   if ( Table("find", "layer", "titlebox") == %badtable) Exit(%ok, "Nothing to update");
  
// Ensure that Layer colour and line style are by layer 
//
   Execute ("menu", "SetStyle",  Table("find", "style", "bylayer")) ;
   Execute ("menu", "SetColor",  Table("find", "color", "bylayer")) ;

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
   		"DB SymbolName = 'patientdetails' OR " +
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
      			hMPD = hSym ;
      			sNewSymName = "mainpatientdetails" ;
      			xyTitleO = xySymO;
      			GetDBData (hSym, "fileno", &sOldFileNo) ;
      			GetDBData (hSym, "units", &sOldUnits) ;
      			// Update data base fields
		   	SetDBData (hMPD, "fileno", sFileNo) ;
		   	SetDBData (hMPD, "age", sAge) ;
		   	SetDBData (hMPD, "diagnosis", sDiagnosis) ;
		   	SetDBData (hMPD, "WorkOrder", sWorkOrder) ;
		   	SetDBData (hMPD, "TemplateEngineer", sTemplateEngineer) ;
		   	SetDBData (hMPD, "orderdate", sOrderDate) ;
		   	SetDBData (hMPD, "patient", sPatient) ;
		   	SetDBData (hMPD, "SEX", sSex) ;
		   	SetDBData (hMPD, "units", sUnits) ;
			}
              		if(StringCompare(sSymName, "patientdetails" )){ 
              			// Replace patientdetails symbol for gloves with text
              			SetData("TextAngle", aAngle);
              			SetData("TextHorzJust", 1);
              			SetData("TextVertJust", 32);
		   	GetDBData (hSym, "Fabric", &sFabric) ;
		   	GetDBData (hSym, "Sleeve", &sSide) ;
 			Execute ("menu", "SetLayer",  Table("find", "layer", "Notes")) ;
      			sText = sSide + "\n " + sPatient + "\n " + sWorkOrder + "\n " + sFabric;
			AddEntity ("text", sText, xySymO);
 			Execute ("menu", "SetLayer",  Table("find", "layer", "Construct")) ;
      			xySymO = CalcXY("relpolar", xySymO, 0.75, aAngle - 90) ;
      			sText = " " + sFileNo + "\n " + sDiagnosis + "\n " + sAge + "\n " + sSex;
			AddEntity ("text", sText, xySymO);
			DeleteEntity(hSym);
		   	SetData("TextAngle", 0.0);
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
			SetDBData(hSym, "fileno", sFileNo) ;
      			}
   	   	}
   	// If an "armarm" symbol is loaded then we need an "armcommon" symbol
   	//
   	if( NewArm) {
   		Import("symbol","armcommon");
   		hSym = AddEntity("symbol", "armcommon", xyTitleO.x , xyTitleO.y + nDeltaTitle); 
   		SetEntityData( hSym, "layer", Table("find", "layer", "Data")) ;
		SetDBData(hSym, "fileno", sFileNo) ;
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

   	
// Find all entities that have used the old file number to label them
// Note how the field ID has to be added for backward compatability
//
   Table("add", "field", "ID", "string");
   sEntSelection = 	"DB fileno = '" + sOldFileNo + "' OR " +
   		"DB ID = '*" + sOldFileNo + "*'" ;	//Note use of WildCard
    
   hChan = Open ("selection", sEntSelection );  
   if (hChan) {
   	ResetSelection (hChan);
	while (hEnt = GetNextSelection(hChan))	{
		// Substitute the New File No directly for Old file No
		sDBData = "" ;
		GetDBData (hEnt, "fileno", &sDBData) ;
		if (StringLength(sDBData) > 0) SetDBData(hEnt, "fileno", sFileNo) ;
		
		// Substitute sFileNo in the ID string
		sDBData = "" ;
		GetDBData (hEnt, "ID", &sDBData) ;
		nLoopTimes = StringLength(sDBData) ;
		if (nLoopTimes  > 0) {
			nOldLen = StringLength(sOldFileNo) ;
			nLoopTimes = nLoopTimes - nOldLen + 1 ;
			nn = 1;
			sNewDBData = "" ;
			while (nn <= nLoopTimes){
				if (StringCompare( sOldFileNo, StringMiddle(sDBData, nn, nOldLen))) {
					sNewDBData = sNewDBData + sFileNo + StringMiddle(sDBData, nn + nOldLen, StringLength(sDBData)) ;
					SetDBData(hEnt, "ID", sNewDBData );
					break ;
					}		
				else
					sNewDBData = sNewDBData  + StringMiddle(sDBData, nn, 1) ;
				nn = nn + 1;
				}
			}		
		}
	Close("selection", hChan);
	}    
	
// Check that units are consistent
//
  if (!StringCompare(sOldUnits, sUnits)){
  	if (Display("message", "YNQuestion", "The units given in the TITLEBOX are different to the units in the Drawing!\nTitlebox = " + sUnits + ", Drawing = " + sOldUnits + "\nUse units from the Drawing?") == %yes)
  		SetDBData (hMPD, "units", sOldUnits) ;
  	}

  		
// Save drawing (but)
// Ensure that on use of CL_CLOSE.D that the .wmf file is created
//
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Export ("drawing", sDocFile, "cad") ;
   SetData("FileChanged", 1);	

   Document("maximize");
   
// End of CL_UPDTE.D ;
//
 
