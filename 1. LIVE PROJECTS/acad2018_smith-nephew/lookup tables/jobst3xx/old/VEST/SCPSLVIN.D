// File Name:	SCPSLVIN.D
// Purpose:	Transfer from the drawn vest details of the 
//		To the SideScoop VB programme
//		SCOOP Sleeves
// Version:	1.01
// Date:	20.Jan.98
//
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

HANDLE
	hChan,
	hCurve,
	hFile,
	hDDE,
	hEnt;

// Body box details
NUMBER
	ii,
	nVertexCount,
	nRadius, 
	nStartAngle, 
	nDeltaAngle;

ANGLE
	aAngle;

STRING
	sTmp,
	sPathJOBST ,
	sClass,
	sAge, 
	sProfileID,
	sData,
	sCurveType,
	sID,
	s1,s2,s3,s4,s5,s6,s7,s8,s9,
	sName ;
	

XY
	xyCenter, 
	xyPt1,
	xyPt2,
	xyScale,
	xyO;
	
// Procedure to write the points of the given curve	
//
 procedure PR_WriteCurve ( ) {
     nVertexCount = GetVertexCount(hCurve) ;
     SetData("UnitLinearType", 0);	// "Inches"
     PrintFile(hFile,  nVertexCount ,"\n") ;
     ii = 1;
     while ( ii <= nVertexCount) {
	GetVertex (hCurve, ii, &xyPt1) ;
	PrintFile(hFile,  xyPt1, "\n") ;
	ii = ii + 1;
	}
 	SetData("UnitLinearType", 6);	// "Inches/Fraction"
} // end of PR_WriteCurve 

// Reset 
    UserSelection ("clear");
    UserSelection("update") ;
    Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));
    Execute ( "menu", "SetColor", Table( "find",  "color", "bylayer"));


// Get user to select Vest
//
   if ( !hEnt = GetEntity ("Select a Sleeve Profile", "all")) Exit(%ok,  "User Cancelled") ; 
   
   GetDBData (hEnt, "Data", &sProfileID);
   GetDBData (hEnt, "curvetype", &sCurveType);
   if (!StringCompare(StringMiddle(sCurveType, 1,6), "sleeve"))Exit(%cancel,  "A sleeve profile was not selected") ;
   
   Display ("cursor", "wait", "Searching for more data");

// Get details from entities and write to file
//
   sTmp =  "DB SymbolName = 'vestbody' OR DB SymbolName = 'mainpatientdetails' OR  DB Data = '" + sProfileID+ "' AND NOT DB curvetype = 'DELETED*' " ; 

   hChan = Open ("selection", sTmp);
   if (hChan ) {
      	ResetSelection(hChan);
	hFile = Open ("file", "C:\\JOBST\\SCOOPNCK.DAT", "write") ;
		SetData("UnitLinearType", 0);	// "Inches"

	PrintFile(hFile, "SLEEVE", "\n");

         	while (hCurve = GetNextSelection(hChan)) {
         	
          	GetEntityClass (hCurve, &sClass) ;
           	
           	sData = "" ; sName = "" ; sID = "" ;
          
          	if (StringCompare("symbol", sClass)) {
   			GetGeometry(hCurve, &sName, &xyO, &xyScale, &aAngle);
   			if (StringCompare("vestbody", sName)) {
				GetDBData (hCurve, "AxillaBackNeckRad", &s1);
				GetDBData (hCurve, "AxillaFrontNeckRad", &s2); 
				GetDBData (hCurve, "ShoulderToBackRaglan", &s3);
				GetDBData (hCurve, "AFNRadRight", &s4); 
				GetDBData (hCurve, "ABNRadRight", &s5); 
				GetDBData (hCurve, "SBRaglanRight", &s6); 
				GetDBData (hCurve, "LtAxillaType", &s7);
				GetDBData (hCurve, "RtAxillaType", &s8); 
				GetDBData (hCurve, "Data", &s9); 
   				}
   			if (StringCompare("mainpatientdetails", sName)) {
				GetDBData (hCurve, "age", &sAge) ;
				}
			}   			
      		else
      			GetDBData (hCurve, "curvetype", &sName) ;
      			
      		if (StringLength (sName) > 0) 
      			PrintFile(hFile, sName, "\n");
      		else
      			PrintFile(hFile, "NoCurveTypeFound\n");
		
		PrintFile(hFile, sClass, "\n") ;
		PrintFile(hFile, UID("get",hCurve),"\n") ;
		
      		GetDBData (hCurve, "Data", &sData) ;
      		if (StringLength (sData) > 0) 
      			PrintFile(hFile, sData, "\n");
      		else
      			PrintFile(hFile, "NoIDFound\n");
      			      		
      		if (StringCompare("marker", sClass)) {
    			GetGeometry(hCurve, &sName, &xyO, &xyScale, &aAngle);     			
      			PrintFile(hFile, xyO, "\n");
      			}
      			      			
      		if (StringCompare("arc", sClass)) {
 			GetGeometry(hCurve, &xyCenter, &nRadius, &nStartAngle, &nDeltaAngle);
			PrintFile(hFile, xyCenter, nRadius, nStartAngle, nDeltaAngle, "\n");
      			}
      			
      		if (StringCompare("line", sClass)) {
 			GetGeometry(hCurve, &xyPt1, &xyPt2);
			PrintFile(hFile, xyPt1, xyPt2, "\n");
      			}
      			
      		if (StringCompare("polyline", sClass) || StringCompare("curve", sClass) ) PR_WriteCurve () ;      	
      			
      		}
		SetData("UnitLinearType", 6);	// "Inches"
   		Close("file", hFile) ;
	}

   Close ( "selection", hChan); 

// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

// Start the application and get DDE link and poke file
//
   Execute ("application", sPathJOBST + "\\VEST\\SCOOPDIA.exe", "normal");
   hDDE = Open ("dde", "scoopdia", "scoopdia");
   if ( hDDE ) {
  	Poke ( hDDE, "txtScoopDataFile",  "C:\\JOBST\\SCOOPNCK.DAT");	
  	Poke ( hDDE, "txtAge",  sAge );	
  	Poke ( hDDE, "txtID",  sProfileID);	
	Poke ( hDDE, "txtAxillaBackNeckRad", s1);
	Poke ( hDDE, "txtAxillaFrontNeckRad", s2); 
	Poke ( hDDE, "txtShoulderToBackRaglan", s3);
	Poke ( hDDE, "txtAFNRadRight", s4); 
	Poke ( hDDE, "txtABNRadRight", s5); 
	Poke ( hDDE, "txtSBRaglanRight", s6); 
	Poke ( hDDE, "txtLtAxillaType", s7);
	Poke ( hDDE, "txtRtAxillaType", s8); 
	Poke ( hDDE, "txtMeshData", s9); 
  	}
   else 
   	Exit (%abort,  "Unable to open \"scoopdia\" dialogue!") ;
   	
   Close ( "dde", hDDE);

 Exit (%ok, "Data Transfer Finished");



