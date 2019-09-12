// File Name:	GL_SEL.D
//
// Purpose:	Starts the VB programme that is used to 
//		draw the CAD Glove
//		Data is based on selecting a glove
//
// Version:	1.01
//
// Date:		29.Sep.95
//
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
// Notes:
// 

Display ("cursor", "wait", "Loading Data");

HANDLE	
	hChan,
	hSym,
	hDDE,
	hEnt;
	
XY	
	xySymO,
	xyScale;
	
ANGLE
	aAngle;
	
STRING	
	sSymName,
	sProfileID,
	sClass,
	sSymSelection,
	sPathJOBST,
	sWorkOrder,
	sSide,
	// Strings are initialized to empty "", this ensures empty strings
	// are poked if there is no DB field.
	s1, s2, s3, s4, s5, s6, s10, s11,s12, s20, s21, s22, s23, s24,s25, s26, s27,s28,s29;	
// Clear user selection
//
   UserSelection ("clear") ;
   UserSelection ("update") ;

// Get profile, identify side 
//
   if (! hCurve=GetEntity ("Select Glove Profile", "all")) Exit (%ok, "User Cancelled") ;

   Display ("cursor", "wait", "Searching for more data");
   GetDBData (hCurve, "ID", &sProfileID);
   nStringLength = StringLength(sProfileID);
   if (StringCompare( "Left", StringMiddle(sProfileID, nStringLength - 3, 4))){
	sSide = "Left";
	}
    if (StringCompare( "Right", StringMiddle(sProfileID, nStringLength - 4, 5))) {
	sSide = "Right";
	}

// Check if this is a valid Glove exit if not
// This is crude check only
//    if (StringLength (sSide) == 0 || !StringCompare(StringMiddle(sProfileID,1,1), "G") ) 
    if (StringLength (sSide) == 0 ) 
    	Exit (%cancel, "A Glove Profile was not selected\n") ;
   
// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

// Start the application and get DDE link
//
   Execute ("application", sPathJOBST + "\\GLOVECAD\\CADGLOVE", "normal");
   hDDE = Open ("dde", "cadglove", "cadglove");
   if ( !hDDE) Exit (%abort, "Unable to open \"cadglove\" dialogue!") ;

// Get all of the Entities for uploading
//
   sSymSelection = 	"DB SymbolName = 'mainpatientdetails' OR " +
   		"(DB SymbolName = 'gloveglove' AND DB Sleeve ='"  + sSide +"') OR " +
   		"(DB ID = '" + sProfileID + "' AND  DB Zipper ='PALMER') OR " +
   		"DB SymbolName = 'glovecommon'"  ;

   nEntityFound = 0 ;
   hChan = Open ("selection", sSymSelection );  
   if (hChan && hDDE ) {
      	ResetSelection(hChan);
      	// Poke side
      	//
      	Poke ( hDDE, "txtSide", sSide);
	// Get Symbol data base fields
	//
          	while (hSym = GetNextSelection(hChan)) {
          		GetEntityClass (hSym, &sClass) ;
          		if (StringCompare("symbol", sClass))
   			GetGeometry(hSym, &sSymName, &xySymO, &xyScale, &aAngle);
      		else
      			GetDBData (hSym, "Zipper", &sSymName) ;
      		if(StringCompare(sSymName, "mainpatientdetails" )){ 
			GetDBData( hSym, "fileno", &s1); Poke ( hDDE, "txtFileNo", s1);
			GetDBData( hSym, "patient", &s2); Poke ( hDDE, "txtPatientName", s2);
 			GetDBData( hSym, "age",  &s3); Poke ( hDDE, "txtAge", s3);
			GetDBData( hSym, "units", &s4); Poke ( hDDE, "txtUnits",s4);
			GetDBData( hSym, "sex", &s5); Poke ( hDDE, "txtSex",s5);
			GetDBData( hSym, "Diagnosis", &s6); Poke ( hDDE, "txtDiagnosis", s6);
			GetDBData( hSym, "WorkOrder", &s10); Poke ( hDDE, "txtWorkOrder",s10);
			Poke ( hDDE, "txtUidMPD", MakeString("long",UID("get",hSym)));
   			}
      		if(StringCompare(sSymName, "glovecommon" )){ 
			GetDBData( hSym, "Fabric", &s11); Poke ( hDDE, "txtFabric",s11);
			Poke ( hDDE, "txtUidGC", MakeString("long",UID("get",hSym)));
   			}
      		if(StringCompare(sSymName, "gloveglove" )){ 
			Poke ( hDDE, "txtUidGlove", MakeString("long",UID("get",hSym)));
   			}
      		if(StringCompare(sSymName, "PALMER" )){ 
			GetDBData( hSym, "TapeLengths", &s20); Poke ( hDDE, "txtTapeLengths",s20);
			GetDBData( hSym, "TapeLengths2", &s21); Poke ( hDDE, "txtTapeLengths2",s21);			
			GetDBData( hSym, "TapeLengthPt1", &s22); Poke ( hDDE, "txtTapeLengthPt1",s22);			
			GetDBData( hSym, "Grams", &s23); Poke ( hDDE, "txtGrams",s23);			
			GetDBData( hSym, "TapeMMs", &s24); Poke ( hDDE, "txtTapeMMs",s24);			
			GetDBData( hSym, "Reduction", &s25); Poke ( hDDE, "txtReduction",s25);			
			GetDBData( hSym, "Data", &s26); Poke ( hDDE, "txtDataGlove",s26);			
			GetDBData( hSym, "WristPleat", &s27); Poke ( hDDE, "txtWristPleat",s27);			
			GetDBData( hSym, "ShoulderPleat", &s28); Poke ( hDDE, "txtShoulderPleat",s28);			
			GetDBData( hSym, "Flap", &s29); Poke ( hDDE, "txtFlap",s29);			
			Poke ( hDDE, "txtUidGlove", MakeString("long",UID("get",hSym)));
   			}
   		}
	}

   Close ( "selection", hChan); 
   Close ( "dde", hDDE);
 
   Exit (%ok, "Patient and Glove Details transfered");

