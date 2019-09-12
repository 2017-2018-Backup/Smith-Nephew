// File Name:	SLV_SEL.D
// Purpose: 	Get Patient Details From Drawing
//		And Pass Control To Visual Basic
//		For Arm Details
// Version:	1.1
// Date:		19.Apr.95
//		G. George
HANDLE	
	hSel,		hBody,		//gg
	hBsel,		hCurv,
	hBchan,		hOrigin,
	hDDE,
	hChan;
XY	
	xyO, xyTmp,
	xyPt1;
STRING
	sFileNo,		sData,		sOtemplateID ,
	sTmp,		sSymbolName,	sPatient,		
	sSEX,		sDiagnosis,	sAge,
	sUnits,		sUID,		sPath,
	sProfileID,	sStyle,		sSleeve,
	sType,		sFileNo,		sPathJOBST ; 
	
NUMBER	
	aAngle;
	
// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

// Get sleeve profile, identify SLEEVE and FileNo
//
   hCurv = GetEntity ("Select a Arm Profile", "curve") ;
   if (hCurv == %zero)Exit(%ok,"No Arm Profile selected");

   Display ("cursor", "wait", "Searching for more data");
   GetDBData (hCurv, "ID", &sProfileID);
   nStringLength = StringLength(sProfileID);
   if (StringCompare( "LeftProfile", StringMiddle(sProfileID, nStringLength - 10, 11))){
	sStyle = StringMiddle(sProfileID, 1, nStringLength - 11);
	sSleeve = "Left";
	sFileNo =  StringMiddle(sProfileID, 5, nStringLength - 15);
	}
    if (StringCompare( "RightProfile",StringMiddle(sProfileID, nStringLength - 11, 12))) {
	sStyle = StringMiddle(sProfileID, 1, nStringLength - 12);
	sSleeve = "Right";
	sFileNo =  StringMiddle(sProfileID, 5, nStringLength - 16);
	}

// Check if this is a valid Arm Curve exit if not
//
    if (StringLength (sSleeve) == 0) Exit (%cancel, "An ARM Profile was not selected\n") ;

// Locate origin marker for selected style
//
   sOtemplateID = sStyle + sSleeve +  "originmark";
   sTmp = "type = 'Marker' AND DB ID = '" + sOtemplateID + "'";
   hChan = Open ("selection", sTmp );
   nMarkersFound = 0;
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
		GetGeometry(hEnt, &sType, &xyPt1, &xyTmp, &aAngle) ;
	                if ( xyPt1.x != 0 || xyPt1.y != 0) {
			if (StringCompare(sTmp, sOtemplateID)) {
				nMarkersFound = nMarkersFound + 1;
				hOrigin = hEnt;
				}
			}
		}
   	}
   Close("selection", hChan); 

// Check if that the markers have been found, otherwise exit
//
    if (nMarkersFound < 1) Exit (%cancel, "Missing origin marker for selected ARM, data not found!\n") ;
    if (nMarkersFound > 1) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;


Execute ("application", sPathJOBST + "\\arm\\armdia", "normal");
	hDDE = Open ( "dde",  "armdia", "armdia");

if (!hDDE)
	Exit( %cancel, "Can't start a DDE Link - Retry");
else {
	Poke(hDDE, "txtSleeve",sSleeve);
	Poke(hDDE, "txtType", "SLEEVE");
	}

// get data from origin marker and poke to vb
// If "titlebox" already exists use data from it

   hChan = Open ("selection", "DB SymbolName = 'mainpatientdetails' AND DB fileno = '" + sFileNo + "'");
   ResetSelection(hChan);
   hSel = GetNextSelection(hChan);
  
   if (hSel) {
	GetDBData( hSel, "patient", &sTmp);
	Poke(hDDE, "txtPatientName", sTmp);
	
   	GetDBData( hSel, "fileno", &sTmp);
	Poke(hDDE, "txtFileNo", sTmp);
	
	GetDBData( hSel, "units", &sTmp);
	Poke(hDDE, "txtinchflag", sTmp);
	
	GetDBData( hSel, "sex", &sTmp);
	Poke(hDDE, "txtSEX", sTmp);
	
	GetDBData( hSel, "diagnosis", &sTmp);
	Poke(hDDE, "txtDiagnosis", sTmp);
	
	GetDBData( hSel, "age", &sTmp);
	Poke(hDDE, "txtAge", sTmp);
	
	GetDBData( hSel, "WorkOrder", &sTmp);
	Poke(hDDE, "txtWorkOrder", sTmp);
	
	Poke ( hDDE, "txtUidMPD", MakeString("long",UID("get",hSel)));
	
	}
 else	{
	Execute("dde", hDDE, "Cancel");
	Exit(%cancel, "Can't find a TitleBox") ; 
	}

  Close("selection",hChan);


// Get Length data from sleevebox and poke to vb
//

   sTmp = "DB SymbolName = 'vestarm'  AND DB Sleeve = '" + sSleeve +  "' AND DB fileno = '" + sFileNo +  "'" ;
   hBchan = Open ("selection", sTmp);
   ResetSelection(hBchan);
   hBsel = GetNextSelection(hBchan);
   if (hBsel) {
	GetDBData( hBsel, "TapeLengths",&sTmp);
	Poke(hDDE, "txtTapeLent",sTmp);	
	}
   else	{
   	Execute("dde", hDDE, "Cancel");
	Exit(%cancel, "Can't find " + sSleeve + " VESTARM symbol to get data from!") ; 
	}
		
   Close("selection",hBchan);


// get the data from the origin and poke to VB
//	
   if (hOrigin) {
	GetDBData( hOrigin, "ID",&sTmp);
	Poke(hDDE, "txtID",sTmp);
	
	GetDBData( hOrigin, "Fabric",&sTmp);
	Poke(hDDE, "txtFabric",sTmp);
	
	GetDBData( hOrigin, "WristPleat",&sTmp);
	Poke(hDDE, "txtWristPleat1",sTmp);
	
	GetDBData( hOrigin, "ShoulderPleat",&sTmp);
	Poke(hDDE, "txtShoulderPleat1",sTmp);
	
	GetDBData( hOrigin, "Contracture",&sTmp);
	Poke(hDDE, "txtContractureWidth",sTmp);

	GetDBData( hOrigin, "Lining",&sTmp);
	Poke(hDDE, "txtLining",sTmp);

	GetDBData( hOrigin, "MM",&sTmp);
	Poke(hDDE, "txtMM",sTmp);
	
	GetDBData( hOrigin, "Stump",&sTmp);
	Poke(hDDE, "txtStump",sTmp);
	
// N.B. Lengths from arm box 
//	GetDBData( hOrigin, "TapeLengths",&sTmp);
//	Poke(hDDE, "txtTapeLent",sTmp);
	
	GetDBData( hOrigin, "TapeMMs",&sTmp);
	Poke(hDDE, "txtTapeMM",sTmp);
	
	GetDBData( hOrigin, "Reduction",&sTmp);
	Poke(hDDE, "txtReduction",sTmp);
	
	GetDBData( hOrigin, "Grams",&sTmp);
	Poke(hDDE, "txtWeight",sTmp);

	// Flap,		
	sTmp = "";
	GetDBData( hOrigin, "Flap",&sTmp);
	Poke(hDDE, "txtFlap",sTmp);
	
	// Gauntlet Multiple fields
	sTmp = "";
	GetDBData( hOrigin, "Gauntlet",&sTmp);
	Poke(hDDE, "txtGauntlet",sTmp);	
	}
  	
	
   Close ( "dde", hDDE);

