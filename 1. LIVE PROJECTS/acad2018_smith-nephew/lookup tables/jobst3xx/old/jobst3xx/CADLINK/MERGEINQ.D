// File Name:	MERGEINQ.D
//
// Purpose:		To display data for a selected symbol 
//
// Version:		1.00
//
// Date:		13.Feb.98
//
// Author:		Gary George
//			Copyright C-Gem Limited, 1988
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
	hMPD,
	hEnt;
	
XY	
	xySymO,
	xyScale;
	
ANGLE
	aAngle;

STRING	
	sSymName,
	sSymSelection, sDialogue, sApplication, sLinkTopic,
	sPathJOBST,
	// Strings are initialized to empty "", this ensures empty strings
	// are poked if there is no DB field.
	s1, s2, s3, s4, s5, s6, s7, s8, s9,
	s10, s11, s12, s13, s14, s15, s16, s17, s18, s19,
	s20, s21, s22, s23, s24, s25, s26, s27, s28, s29, 
	s30, s31, s32, s33, s34, s35, s36, s37, s38, s39,
	s40, s41, s42, s43, s44, s45, s46, s47, s48, s49;
	
// Procedure to start the application and get DDE link
//
procedure PR_OpenDDE () {
   Execute ("application",  sPathJOBST + sApplication , "normal");
   if (StringLength(sLinkTopic) ==0) sLinkTopic = sDialogue ;
   
   hDDE = Open ( "dde",  sDialogue , sLinkTopic );
   
   if ( !hDDE) Exit (%abort, "Unable to open \"" + sDialogue + "\"dialogue!") ;
   
   GetDBData( hMPD, "fileno", &s1); Poke ( hDDE, "txtFileNo", s1);
   GetDBData( hMPD, "patient", &s2); Poke ( hDDE, "txtPatientName", s2);
   GetDBData( hMPD, "age",  &s3); Poke ( hDDE, "txtAge", s3);
   GetDBData( hMPD, "units", &s4); Poke ( hDDE, "txtinchflag",s4);  Poke ( hDDE, "txtUnits",s4);
   GetDBData( hMPD, "sex", &s5); Poke ( hDDE, "txtSEX",s5); Poke ( hDDE, "txtSex",s5);
   GetDBData( hMPD, "Diagnosis", &s6); Poke ( hDDE, "txtDiagnosis", s6);
   GetDBData( hMPD, "WorkOrder", &s7); Poke ( hDDE, "txtWorkOrder",s7);
   Poke ( hDDE, "txtUidMPD", MakeString("long",UID("get",hMPD)));
   
} // end procedure 

// procedures to save duplication of code.
//
procedure PR_PokeVestFields () {
	GetDBData( hSym, "LtSCir", &s10);	Poke ( hDDE, "txtCir(0)", s10);
	GetDBData( hSym, "RtSCir", &s11);	Poke ( hDDE, "txtCir(1)", s11);
	GetDBData( hSym, "NeckCir", &s12);	Poke ( hDDE, "txtCir(2)", s12);
	GetDBData( hSym, "SWidth", &s13);	Poke ( hDDE, "txtCir(3)", s13);	  
	GetDBData( hSym, "S_Waist", &s14);	Poke ( hDDE, "txtCir(4)", s14);
	GetDBData( hSym, "ChestCir", &s15);	Poke ( hDDE, "txtCir(5)", s15);
	GetDBData( hSym, "WaistCir", &s16 );	Poke ( hDDE, "txtCir(6)", s16);
	GetDBData( hSym, "S_Breast", &s19);	Poke ( hDDE, "txtCir(9)", s19);
	GetDBData( hSym, "BreastCir", &s20);	Poke ( hDDE, "txtCir(10)", s20);
	GetDBData( hSym, "NippleCir", &s21);	Poke ( hDDE, "txtCir(11)", s21);
	GetDBData( hSym, "Fabric", &s22); 		Poke ( hDDE, "txtCombo(10)", s22);
	GetDBData( hSym, "BraLtCup", &s23);	Poke ( hDDE, "txtCombo(3)", s23);
	GetDBData( hSym, "BraRtCup", &s24);	Poke ( hDDE, "txtCombo(4)", s24);
	GetDBData( hSym, "BraRtDisk", &s25);	Poke ( hDDE, "txtRightDisk", s25);
	GetDBData( hSym, "BraLtDisk", &s26);	Poke ( hDDE, "txtLeftDisk", s26);
	GetDBData( hSym, "LtAxillaType", &s27);	Poke ( hDDE, "txtCombo(5)", s27);
	GetDBData( hSym, "RtAxillaType", &s28);	Poke ( hDDE, "txtCombo(6)", s28);
	GetDBData( hSym, "Closure", &s29);		Poke ( hDDE, "txtCombo(9)", s29);
	GetDBData( hSym, "ChestCirUserFac", &s31);	Poke ( hDDE, "txtCombo(1)", s31);
	GetDBData( hSym, "WaistCirUserFac", &s30);	Poke ( hDDE, "txtCombo(0)", s30);
	GetDBData( hSym, "BreastCirUserFac", &s32);	Poke ( hDDE, "txtCombo(2)", s32);
	GetDBData( hSym, "NeckType", &s33);		Poke ( hDDE, "txtCombo(7)", s33);
	GetDBData( hSym, "BackNeckType", &s34);	Poke ( hDDE, "txtCombo(8)", s34);
	GetDBData( hSym, "NeckDimension", &s35);  	Poke ( hDDE, "txtFrontNeck", s35);
	GetDBData( hSym, "BackNeckDim", &s36);	Poke ( hDDE, "txtBackNeck", s36);
	
} // End procedure

procedure PR_PokeArmFields () {
	GetDBData( hSym, "sleeve", &s20) ; 	Poke(hDDE, "txtSleeve",s20);
	GetDBData( hSym, "ID",&s21); 		Poke(hDDE, "txtID",s21);
	GetDBData( hSym, "WristPleat",&s22);	Poke(hDDE, "txtWristPleat1",s22);
	GetDBData( hSym, "ShoulderPleat",&s23); 	Poke(hDDE, "txtShoulderPleat1",s23);
	GetDBData( hSym, "Contracture",&s24); 	Poke(hDDE, "txtContractureWidth",s24);
	GetDBData( hSym, "Lining",&s25); 		Poke(hDDE, "txtLining",s25);
	GetDBData( hSym, "MM",&s26); 		Poke(hDDE, "txtMM",s26);
	GetDBData( hSym, "Stump",&s27); 		Poke(hDDE, "txtStump",s27);
	GetDBData( hSym, "TapeLengths",&s28); Poke(hDDE, "txtTapeLent",s28);
	GetDBData( hSym, "TapeMMs",&s29); Poke(hDDE, "txtTapeMM",s29);
	GetDBData( hSym, "Reduction",&s30); Poke(hDDE, "txtReduction",s30);
	GetDBData( hSym, "Grams",&s31); Poke(hDDE, "txtWeight",s31);
	GetDBData( hSym, "Flap",&s32); Poke(hDDE, "txtFlap",s32);
	GetDBData( hSym, "Gauntlet",&s33); Poke(hDDE, "txtGauntlet",s33);	
	Poke ( hDDE, "txtUidArm", MakeString("long",UID("get",hSym)));    	

} // End procedure

procedure PR_PokeLegFields () {

	GetDBData( hSym, "Leg", &s27);  Poke ( hDDE, "txtLeg", s27);
   	GetDBData( hSym, "ToeStyle", &s13);	Poke ( hDDE, "txtToeStyle", s13);
   	GetDBData( hSym, "FootPleat1", &s14);	Poke ( hDDE, "txtFootPleat1", s14);
   	GetDBData( hSym, "FootPleat2", &s15);	Poke ( hDDE, "txtFootPleat2", s15);
   	GetDBData( hSym, "TopLegPleat1", &s16);	Poke ( hDDE, "txtTopLegPleat1", s16);
   	GetDBData( hSym, "TopLegPleat2", &s17);	Poke ( hDDE, "txtTopLegPleat2", s17);
   	GetDBData( hSym, "FootLength", &s18);	Poke ( hDDE, "txtFootLength", s18);

} // End procedure

// Clear user selection
//
   UserSelection ("clear") ;
   UserSelection ("update") ;  

// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

// Get the mainpatient details of the symbols for uploading
//
   sSymSelection = "DB SymbolName = 'mainpatientdetails' " ;
   hChan = Open ("selection", sSymSelection );  
   if (hChan) {
      	ResetSelection(hChan);
          while (hSym = GetNextSelection(hChan)) {
   		GetGeometry(hSym, &sSymName, &xySymO, &xyScale, &aAngle);
		if(StringCompare(sSymName, "mainpatientdetails" )) hMPD = hSym ;
   		}
   }
   Close ( "selection", hChan); 

// Check we have a MPD
//
   if (!hMPD) Exit (%cancel, "No patient details found") ;
   
// Prompt the user to select a symbol
//
   if (!hSym =GetEntity ("Select symbol", "symbol")) Exit (%ok, "User Cancelled") ;
  
   Display ("cursor", "wait", "Loading Data");
   GetGeometry(hSym, &sSymName, &xySymO, &xyScale, &aAngle);
Print ("\n", sSymName) ;
   
   if(StringCompare(sSymName, "armarm" )){
     	sApplication = "\\ARM\\ARMDIA" ;
	sDialogue = "armdia";
	PR_OpenDDE() ;
  	PR_PokeArmFields() ;
 	Poke(hDDE, "txtType", "ARM");
   	}
   	
   if(StringCompare(sSymName, "vestarm" )){
     	sApplication = "\\ARM\\ARMDIA" ;
	sDialogue = "armdia";
	PR_OpenDDE() ;
 	PR_PokeArmFields() ;
	Poke(hDDE, "txtType", "SLEEVE");
   	}
   	
   if(StringCompare(sSymName, "vestbody" )){ 
     	sApplication =  "\\VEST\\VESTDIA" ;
	sDialogue = "vestdia";
	
	PR_OpenDDE() ;
	PR_PokeVestFields() ;
	Poke ( hDDE, "txtUidVB", MakeString("long",UID("get",hSym)));
	
   	}
   	
   if(StringCompare(sSymName, "waistbody" )){ 
     	sApplication =  "\\waist\\whboddia" ;
	sDialogue = "whboddia";
	
	PR_OpenDDE() ;
	
   	GetDBData( hSym, "TOSCir", &s10); 		Poke ( hDDE, "txtTOSCir", s10);
   	GetDBData( hSym, "TOSGivenRed", &s11); 	Poke ( hDDE, "txtTOSRed", s11);
   	GetDBData( hSym, "TOSHt", &s12);		Poke ( hDDE, "txtTOSHt", s12);
   	GetDBData( hSym, "WaistCir", &s13);		Poke ( hDDE, "txtWaistCir", s13);
   	GetDBData( hSym, "WaistGivenRed", &s14);	Poke ( hDDE, "txtWaistRed", s14);
   	GetDBData( hSym, "WaistHt", &s15);		Poke ( hDDE, "txtWaistHt", s15);
   	GetDBData( hSym, "MidPointCir", &s16);	Poke ( hDDE, "txtMidPointCir", s16);
   	GetDBData( hSym, "MidPointGivenRed", &s17);	Poke ( hDDE, "txtMidPointRed", s17);
   	GetDBData( hSym, "MidPointHt", &s18);	Poke ( hDDE, "txtMidPointHt", s18);
   	GetDBData( hSym, "LargestCir", &s20);	Poke ( hDDE, "txtLargestCir", s20);
   	GetDBData( hSym, "LargestGivenRed", &s21);	Poke ( hDDE, "txtLargestRed", s21);
   	GetDBData( hSym, "LargestHt", &s22);		Poke ( hDDE, "txtLargestHt", s22);
   	GetDBData( hSym, "LeftThighCir", &s23);	Poke ( hDDE, "txtLeftThighCir", s23);
   	GetDBData( hSym, "RightThighCir", &s24);	Poke ( hDDE, "txtRightThighCir", s24);
   	GetDBData( hSym, "ThighGivenRed", &s25);	Poke ( hDDE, "txtThighRed", s25);
   	GetDBData( hSym, "FoldHt", &s26);		Poke ( hDDE, "txtFoldHt", s26);
   	GetDBData( hSym, "CrotchStyle", &s27);	Poke ( hDDE, "txtCrotchStyle", s27);
   	GetDBData( hSym, "ID", &s28);			Poke ( hDDE, "txtLegStyle", s28);
   	GetDBData( hSym, "Body", &s29);		Poke ( hDDE, "txtBody", s29);
	Poke ( hDDE, "txtUidWHBody", MakeString("long",UID("get",hSym)));
	Poke ( hDDE, "txtUidTitle", MakeString("long",UID("get",hMPD)));

   	}

   if(StringCompare(sSymName, "suitbody" )){ 
     	sApplication =   "\\BODY\\BODYSUIT" ;
	sDialogue = "bodysuit";
	
	PR_OpenDDE() ;
	PR_PokeVestFields() ;	

	// Extra bodysuit specific Fields
	GetDBData( hSym, "SFButt", &s40);	Poke ( hDDE, "txtCir(12)", s40);
	GetDBData( hSym, "SLgButt", &s41);	Poke ( hDDE, "txtCir(13)", s41);
	GetDBData( hSym, "LgButtCir", &s42);	Poke ( hDDE, "txtCir(14)", s42);
	GetDBData( hSym, "LgButtCirUserFac", &s47);	Poke ( hDDE, "txtCombo(11)", s47);
	GetDBData( hSym, "LtThCir", &s43);	Poke ( hDDE, "txtCir(15)", s43);
	GetDBData( hSym, "ThCirUserFac", &s48);	Poke ( hDDE, "txtCombo(12)", s48);
	GetDBData( hSym, "RtThCir", &s44);	Poke ( hDDE, "txtCir(16)", s44);
	GetDBData( hSym, "LegStyle", &s45);	Poke ( hDDE, "txtLegStyle", s45);
	GetDBData( hSym, "CrotchStyle", &s46); Poke ( hDDE, "txtCrotchStyle", s46);

	Poke ( hDDE, "txtUidVB", MakeString("long",UID("get",hSym)));
	}
  
   if(StringCompare(sSymName, "headneck" )){ 
     	sApplication =   "\\HEADNECK\\HeadNeck";
	sDialogue = "HeadNeck";
	sLinkTopic = "Form1" ;
	PR_OpenDDE() ;
	
   	GetData("DrawingFile", &s10);   Poke(hDDE, "txtDraw", s10);	
	
	GetDBData( hSym, "UID",&s11);	Poke(hDDE, "txtUID",s11);
	GetDBData( hSym, "HeadNeck", &s12);	Poke(hDDE, "txtMeasurements", s12);
   	GetDBData( hSym, "Fabric", &s13);	Poke(hDDE, "txtFabric", s13);
	GetDBData( hSym, "Data", &s14);	Poke(hDDE, "txtData", s14);
	
	Poke ( hDDE, "txtUidHN", MakeString("long",UID("get",hSym)));
	}

   if(StringCompare(sSymName, "legleg" )){ 
     	sApplication =   "\\leg\\lglegdia";
	sDialogue =  "lglegdia";
	PR_OpenDDE() ;
	PR_PokeLegFields () ;
	GetDBData( hSym, "TapeLengthsPt1", &s10);
  	GetDBData( hSym, "TapeLengthsPt2", &s11);
		s10 = s10 + s11 ;
	Poke ( hDDE, "txtLeftLengths", s10 );
	GetDBData( hSym, "AnkleTape", &s12);	Poke ( hDDE, "txtLeftAnkle", s12);
   	GetDBData( hSym, "Pressure", &s19);	Poke ( hDDE, "txtLeftTemplate", s19);
   	GetDBData( hSym, "Fabric", &s20);		Poke ( hDDE, "txtFabric", s20);
   	GetDBData( hSym, "Anklet", &s21);		Poke ( hDDE, "txtAnklet", s21);
   	GetDBData( hSym, "ThighLength", &s22);	Poke ( hDDE, "txtThighLength", s22);
   	GetDBData( hSym, "KneeLength", &s23);	Poke ( hDDE, "txtKneeLength", s23);
   	GetDBData( hSym, "ThighBand", &s24);	Poke ( hDDE, "txtThighBandAK", s24);
    	GetDBData( hSym, "ThighBandBK", &s25);	Poke ( hDDE, "txtThighBandBK", s25);
  	GetDBData( hSym, "KneeBand", &s26);	Poke ( hDDE, "txtKneeBand", s26);
	Poke ( hDDE, "txtChosenStyle", "-1");
	
	Poke ( hDDE, "txtUidLeftLeg", MakeString("long",UID("get",hSym)));
 	}

     if(StringCompare(sSymName, "waistleg" )){ 
     	sApplication =   "\\waist\\whlegdia";
	sDialogue =  "whlegdia";
	PR_OpenDDE() ;
	PR_PokeLegFields () ;
	GetDBData( hSym, "TapeLengthsPt1", &s10);
  	GetDBData( hSym, "TapeLengthsPt2", &s11);
		s10 = s10 + s11 ;
	Poke ( hDDE, "txtTapeLengths", s10 );
  	GetDBData( hSym, "AnkleTape", &s12); 	Poke ( hDDE, "txtAnkleTape", s12);
   	GetDBData( hSym,"Data", &s30);	   	Poke ( hDDE, "txtData", s30);
	Poke ( hDDE, "txtUidLeg", MakeString("long",UID("get",hSym))) ;
 	}
 	
 	
    if(StringCompare(sSymName, "gloveglove" )){  
     	sApplication =   "\\GLOVEMAN\\MANGLOVE";
	sDialogue =  "manglove";
	PR_OpenDDE() ;
         	GetDBData( hSym, "Sleeve", &s10);			Poke ( hDDE, "txtSide", s10);
	GetDBData( hSym, "TapeLengths", &s20); 		Poke ( hDDE, "txtTapeLengths",s20);
	GetDBData( hSym, "TapeLengths2", &s21);  	Poke ( hDDE, "txtTapeLengths2",s21);	
	GetDBData( hSym, "TapeLengthPt1", &s22); 	Poke ( hDDE, "txtTapeLengthPt1",s22);			
	GetDBData( hSym, "Grams", &s23); 			Poke ( hDDE, "txtGrams",s23);			
	GetDBData( hSym, "TapeMMs", &s24); 		Poke ( hDDE, "txtTapeMMs",s24);			
	GetDBData( hSym, "Reduction", &s25); 		Poke ( hDDE, "txtReduction",s25);			
	GetDBData( hSym, "Data", &s26);			Poke ( hDDE, "txtDataGlove",s26);			
	GetDBData( hSym, "WristPleat", &s27); 		Poke ( hDDE, "txtWristPleat",s27);			
	GetDBData( hSym, "ShoulderPleat", &s28); 	Poke ( hDDE, "txtShoulderPleat",s28);			
	GetDBData( hSym, "Flap", &s29); 			Poke ( hDDE, "txtFlap",s29);			
	Poke ( hDDE, "txtUidGlove", MakeString("long",UID("get",hSym)));
	}


      Close ( "dde", hDDE);