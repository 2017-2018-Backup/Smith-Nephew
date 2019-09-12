// File Name:	ARM_IN.D
// Purpose: 	Get Patient Details From Drawing
//		And Pass Control To Visual Basic
//		For Arm Details
//		Called by the shells ARML.D and ARMR.D
// Version:	3
// Date:	19/July/94
// By:		Ciaran McKavanagh
//		G. George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 13.Jan.98		GG	Added MeshData
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
	sSymSelection,
	sPathJOBST,
	sWorkOrder,
	sData,
	// Strings are initialized to empty "", this ensures empty strings
	// are poked if there is no DB field.
	s1, s2, s3, s4, s5, s6, s10, s11, s12, s13, s14, s15, s16, s17,  s20, s21, s22, s23, s24, s25, s26, s27,
	s28, s29, s30, s31, s32, s33, s34, s35;	

//STRING sSleeve; Set by calling shell

   
// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

// Start the application and get DDE link
//
   Execute ("application",  sPathJOBST + "\\arm\\armdia", "normal");
   hDDE = Open ( "dde",  "armdia", "armdia");
   if ( !hDDE) Exit (%abort, "Unable to open \"armdia\" dialogue!") ;

// Get all of the symbols for uploading
//
   sSymSelection = 	"DB SymbolName =  'mainpatientdetails' OR " +
   		"(DB SymbolName = 'vestarm' AND DB Sleeve ='"  + sSleeve +"') OR " +
   		"DB SymbolName =  'vestbody'"  ;

   hChan = Open ("selection", sSymSelection );  
   if (hChan && hDDE ) {
      	ResetSelection(hChan);
      	// Poke side
	Poke(hDDE, "txtSleeve",sSleeve);
	Poke(hDDE, "txtType", "SLEEVE") ;
          	while (hSym = GetNextSelection(hChan)) {
   		GetGeometry(hSym, &sSymName, &xySymO, &xyScale, &aAngle);
      		if(StringCompare(sSymName, "mainpatientdetails" )){ 
			GetDBData( hSym, "fileno", &s1); Poke ( hDDE, "txtFileNo", s1);
			GetDBData( hSym, "patient", &s2); Poke ( hDDE, "txtPatientName", s2);
 			GetDBData( hSym, "age",  &s3); Poke ( hDDE, "txtAge", s3);
			GetDBData( hSym, "units", &s4); Poke ( hDDE, "txtinchflag",s4);
			GetDBData( hSym, "sex", &s5); Poke ( hDDE, "txtSEX",s5);
			GetDBData( hSym, "Diagnosis", &s6); Poke ( hDDE, "txtDiagnosis", s6);
			GetDBData( hSym, "WorkOrder", &s10); Poke ( hDDE, "txtWorkOrder",s10);
			Poke ( hDDE, "txtUidMPD", MakeString("long",UID("get",hSym)));
   			}

            	if(StringCompare(sSymName, "vestbody" )){ 
			GetDBData( hSym, "Fabric", &s11); Poke ( hDDE, "txtFabric",s11);
			Poke ( hDDE, "txtUidAC", MakeString("long",UID("get",hSym)));
			
			//Carrys mesh data for left and right arms	
			GetDBData( hSym, "Data", &s16); Poke ( hDDE, "txtMeshData",s16);
			
			//Get the ID of the draw vest
			//Use this to link the drawn sleeve raglan and the vest raglan
			// for use with side scoops
			GetDBData( hSym, "ID", &s17); Poke ( hDDE, "txtVestID",s17);
		
		   	if (StringCompare(sSleeve, "Right")) {
			// If Left and Right axilla are of different heights then there is a specific
			// set of values for the Right.  If these are found they are used.
			//
				GetDBData ( hSym,  "AFNRadRight", &s12) ;
				if ( !Value ( "scalar", s12 )) GetDBData ( hSym,  "AxillaFrontNeckRad", &s12) ;
				sData = s12 ;
		
				GetDBData ( hSym,  "ABNRadRight", &s13) ;
				if ( !Value ( "scalar", s13)) GetDBData ( hSym,  "AxillaBackNeckRad", &s13) ;
				sData = sData + " " + s13;		
		
				GetDBData ( hSym,  "SBRaglanRight", &s14) ;
				if ( !Value ( "scalar", s14)) GetDBData ( hSym,  "ShoulderToBackRaglan", &s14) ;
				sData = sData + " " + s14 ;		
		
				GetDBData( hSym,  "RtAxillaType", &s15);
				sData = s15 + " " + sData  ;		
				}
			else {
				GetDBData ( hSym,  "AxillaFrontNeckRad", &s12) ;
				sData = s12;

    				GetDBData ( hSym,  "AxillaBackNeckRad", &s13) ;
				sData = sData + " " + s13;		

    				GetDBData ( hSym,  "ShoulderToBackRaglan", &s14) ;
				sData = sData + " " + s14 ;		

				GetDBData( hSym,  "LtAxillaType", &s15);
				sData = s15 + " " + sData  ;		
				}
   			Poke ( hDDE, "txtVestRaglan", sData) ;  
   			}

         		if(StringCompare(sSymName, "vestarm" )){ 
			GetDBData( hSym, "ID",&s21); Poke(hDDE, "txtID",s21);
			GetDBData( hSym, "WristPleat",&s22);	Poke(hDDE, "txtWristPleat1",s22);
			GetDBData( hSym, "ShoulderPleat",&s23); Poke(hDDE, "txtShoulderPleat1",s23);
			GetDBData( hSym, "Contracture",&s24); Poke(hDDE, "txtContractureWidth",s24);
			GetDBData( hSym, "Lining",&s25); Poke(hDDE, "txtLining",s25);
			GetDBData( hSym, "MM",&s26); Poke(hDDE, "txtMM",s26);
			GetDBData( hSym, "Stump",&s27); Poke(hDDE, "txtStump",s27);
			GetDBData( hSym, "TapeLengths",&s28); Poke(hDDE, "txtTapeLent",s28);
			GetDBData( hSym, "TapeMMs",&s29); Poke(hDDE, "txtTapeMM",s29);
			GetDBData( hSym, "Reduction",&s30); Poke(hDDE, "txtReduction",s30);
			GetDBData( hSym, "Grams",&s31); Poke(hDDE, "txtWeight",s31);
			GetDBData( hSym, "Flap",&s32); Poke(hDDE, "txtFlap",s32);
			GetDBData( hSym, "Gauntlet",&s33); Poke(hDDE, "txtGauntlet",s33);	
			Poke ( hDDE, "txtUidArm", MakeString("long",UID("get",hSym)));
   			}
   		}
	}
   
   Close ( "selection", hChan); 
   Close ( "dde", hDDE);
 
   Exit (%ok, "Patient and Vest Arm details transfered");

