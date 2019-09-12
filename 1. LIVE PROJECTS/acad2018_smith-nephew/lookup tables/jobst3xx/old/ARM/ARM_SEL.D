// File Name:	ARM_SEL.D
// Purpose: 	Get Patient Details From Drawing
//		And Pass Control To Visual Basic
//		For Arm Details
// Version:	1.1
// Date:		19.Apr.95
//		G. George
HANDLE	
	hSym,		
	hCurv,
	hEnt,
	hOrigin,
	hDDE,
	hChan;
XY	
	xyO, xyTmp, xySymO, xyScale, xyPt1;


STRING	
	sSymName,
	sSymSelection,
	sPathJOBST,
	sWorkOrder,
	sSleeve,
	sPathJOBST,
	sFileNo,
	sData,		
	sOtemplateID ,
	sTmp,		
	sPath,	
	sProfileID,	
	sStyle,
	sType,
	// Strings are initialized to empty "", this ensures empty strings
	// are poked if there is no DB field.
	s1, s2, s3, s4, s5, s6, 
	s10, s11,
	s20, s21, s22,
	s30, s31, s32, s33, s34, s35, s36, s37,	s38, s39, s40, s41, s42;			
	
NUMBER	
	aAngle;
	
// Get sleeve profile, identify ARM and FileNo
//
   hCurv = GetEntity ("Select an Arm Profile", "curve") ;
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

// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "C:\\JOBST", "DRAFIX.INI") ;

// Start the application and get DDE link
//
   Execute ("application", sPathJOBST  + "\\arm\\armdia", "normal");
	hDDE = Open ( "dde",  "armdia", "armdia");

   if (!hDDE)
	Exit( %cancel, "Can't start a DDE Link - Retry");
   else {
	Poke(hDDE, "txtSleeve",sSleeve);
	Poke(hDDE, "txtType", "ARM");
	}
// Get all of the symbols for uploading
//

   sSymSelection = 	"(DB SymbolName = 'mainpatientdetails' AND DB fileno = '" + sFileNo + "') OR" +
   		"(DB SymbolName = 'armarm' AND DB Sleeve ='"  + sSleeve +"' AND DB fileno = '" + sFileNo + "') OR " +
   		"(DB SymbolName = 'armcommon' AND DB fileno = '" + sFileNo + "')"  ;

   hChan = Open ("selection", sSymSelection );  
   if (hChan && hDDE ) {
      	ResetSelection(hChan);
      	// Poke side
	Poke(hDDE, "txtSleeve",sSleeve);
	Poke(hDDE, "txtType", "ARM");
          	while (hSym = GetNextSelection(hChan)) {
   		GetGeometry(hSym, &sSymName, &xySymO, &xyScale, &aAngle);
      		if(StringCompare(sSymName, "mainpatientdetails" )){ 
			GetDBData( hSym, "fileno", &s1); Poke ( hDDE, "txtFileNo", s1);
			GetDBData( hSym, "patient", &s2); Poke ( hDDE, "txtPatientName", s2);
 			GetDBData( hSym, "age",  &s3); Poke ( hDDE, "txtAge", s3);
			GetDBData( hSym, "units", &s4); Poke ( hDDE, "txtinchflag",s4);
			GetDBData( hSym, "sex", &s5); Poke ( hDDE, "txtSEX",s5);
			GetDBData( hSym, "Diagnosis", &s6); Poke ( hDDE, "txtDiagnosis", s6);
			Poke ( hDDE, "txtUidMPD", MakeString("long",UID("get",hSym)));
   			}
      		if(StringCompare(sSymName, "armcommon" )){ 
			GetDBData( hSym, "WorkOrder", &s10); Poke ( hDDE, "txtACwo",s10);
			GetDBData( hSym, "Fabric", &s11); Poke ( hDDE, "txtFabric",s11);
			Poke ( hDDE, "txtUidAC", MakeString("long",UID("get",hSym)));
   			}
      		if(StringCompare(sSymName, "armarm" )){ 
			GetDBData( hSym, "ID",&s21); Poke(hDDE, "txtID",s21);
			GetDBData( hSym, "TapeLengths",&s22); Poke(hDDE, "txtTapeLent",s22);
			Poke ( hDDE, "txtUidArm", MakeString("long",UID("get",hSym)));
   			}
   		}
	}
   
   Close ( "selection", hChan); 

// get the data from the origin and poke to VB
//	
   if (hOrigin) {
	GetDBData( hOrigin, "ID",&s30); Poke(hDDE, "txtID",s30);
	
	GetDBData( hOrigin, "WristPleat",&s31); Poke(hDDE, "txtWristPleat1",s31);
	
	GetDBData( hOrigin, "ShoulderPleat",&s32); Poke(hDDE, "txtShoulderPleat1",s32);
	
	GetDBData( hOrigin, "Contracture",&s33); Poke(hDDE, "txtContractureWidth",s34);

	GetDBData( hOrigin, "Lining",&s35); Poke(hDDE, "txtLining",s35);

	GetDBData( hOrigin, "MM",&s36); Poke(hDDE, "txtMM",s36);
	
	GetDBData( hOrigin, "Stump",&s37); Poke(hDDE, "txtStump",s37);
	
// N.B. Lengths from arm box 
//	GetDBData( hOrigin, "TapeLengths",&sTmp);
//	Poke(hDDE, "txtTapeLent",sTmp);
	
	GetDBData( hOrigin, "TapeMMs",&s38); Poke(hDDE, "txtTapeMM",s38);
	
	GetDBData( hOrigin, "Reduction",&s39); Poke(hDDE, "txtReduction",s39);
	
	GetDBData( hOrigin, "Grams",&s40); Poke(hDDE, "txtWeight",s40);

	// Flap,		
	GetDBData( hOrigin, "Flap",&s41); Poke(hDDE, "txtFlap",s41);
	
	// Gauntlet Multiple fields
	GetDBData( hOrigin, "Gauntlet",&s42);	Poke(hDDE, "txtGauntlet",s42);	
	}
  	
	
   Close ( "dde", hDDE);
 
   Exit (%ok, "Patient and Arm Details transfered");






