// File Name:	SCOOP_IN.D
// Purpose:	Transfer from the drawn vest details of the 
//		To the SideScoop VB programme
//
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
	sProfileID,
	sData,
	sID,
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
     SetData("UnitLinearType", 0);	// "Inches"
     nVertexCount = GetVertexCount(hCurve) ;
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
   if ( !hEnt = GetEntity ("Select Vest Raglan Profile", "all")) Exit(%ok,  "User Cancelled") ;
   
   GetDBData (hEnt, "ID", &sProfileID);
   
   Display ("cursor", "wait", "Searching for more data");

// Get details from entities and write to file
//
   sTmp =  "DB SymbolName = 'vestbody' OR  DB Data = '" +sProfileID + "*' AND NOT DB curvetype = 'DELETED*' " ; 

   hChan = Open ("selection", sTmp);
   if (hChan ) {
      	ResetSelection(hChan);
	hFile = Open ("file", "C:\\JOBST\\SCOOPNCK.DAT", "write") ;
		SetData("UnitLinearType", 0);	// "Inches"

         	while (hCurve = GetNextSelection(hChan)) {
         	
          	GetEntityClass (hCurve, &sClass) ;
           	
           	sData = "" ; sName = "" ; sID = "" ;
          
          	if (StringCompare("symbol", sClass))
   			GetGeometry(hCurve, &sName, &xyO, &xyScale, &aAngle);
      		else
      			GetDBData (hCurve, "curvetype", &sName) ;
      			
      		if (StringLength (sName) > 0) 
      			PrintFile(hFile, sName, "\n");
      		else
      			PrintFile(hFile, "NoCurveTypeFound\n");
		
		PrintFile(hFile, sClass, "\n") ;
		PrintFile(hFile, UID("get",hCurve),"\n") ;

      		GetDBData (hCurve, "ID", &sID) ;
      		if (StringLength (sID) > 0) 
      			PrintFile(hFile, sID, "\n");
      		else
      			PrintFile(hFile, "NoIDFound\n");
      			
      		GetDBData (hCurve, "Data", &sData) ;
      		if (StringLength (sData) > 0) 
      			PrintFile(hFile, sData, "\n");
      		else
      			PrintFile(hFile, "NoDataFound\n");
      			
      			      		
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
		SetData("UnitLinearType", 6);	// "Inches/Fraction"
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
  	Poke ( hDDE, "txtID",  sProfileID );	
  	}
   else 
   	Exit (%abort,  "Unable to open \"scoopdia\" dialogue!") ;
   	
   Close ( "dde", hDDE);

 Exit (%ok, "Data Transfer Finished");



