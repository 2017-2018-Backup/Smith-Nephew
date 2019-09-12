// File Name:	WHLG_ZIP.D
// Purpose:	General Zip tool
//		Lower extremities
//		Waist Heights
// Version:	1.01
// Date:		15.Dec.1994
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
Display ("cursor", "wait", "Loading dialogue");

HANDLE
	hDlg,
	hEnt,
	hCurve,
	hChan,
	hTextEnt,
	hLeg,
	hLayer;

XY
	xySmallestCir,
	xyConstruct,
	xyStart,
	xyEnd,
	xyHorizStart,
	xyHorizEnd,
	xyMax, xyMin,
	xyAnkle,
	xyOtemplate,
	xyPt1,
	xyTmp;

STRING
	sClass,
	sZipperID,
	sElastic,
	sDlgProfileOffset,
	sDlgDistalElasticList, 
	sElasticDistal,
	sDlgProximalElasticList,
	sElasticProximal,
	sOffsetType,
	sOffset,
	sZipType,
	sUnits,
	sZipLength,
	sProfileID,
	sAnkleID,
	sOtemplateID,
	sTmp, sTmp1,
	sLeg,
	sStyleID,
	sDlgElasticList ,
	sDlgLengthList,
	sLine,
	sType;

NUMBER
	nMedial,
	nElastic,
	nType,
	ii, nn,
	nY,nX,
	nEOSRad,
	nZipLength,
	nHeelOff,
	nHeelRad;

BOOL
	TwoPoint,
	PointLength,
	FirstPoint,
	OldDrawing,
	StandardOffset,
	GivenOffset,
	FootLess,
	Loop;

MSGID
	Ok;

   STRING	 sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\WAIST" );
   
   procedure PR_SetZipperID () {
	// Label entity with ID string and make Zipper %true
   	SetDBData(hEnt, "ID", sZipperID );
   	SetDBData(hEnt, "Zipper", "1" );
   }
   
// Start Dialogue
// Create Dialog
//
   sUnits = "Inches";
   
   Loop = %true ;
   nMinZipLength = 5 ;
   
   sDlgElasticList =  "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
   sDlgProximalElasticList =  "No Elastic\n" + sDlgElasticList;
   sDlgDistalElasticList =  "No Elastic\n" + sDlgElasticList;

   sDlgLengthList =  "Two Points\nPoint & Length";   
   sDlgProfileOffset = "1-1/8\"\nFirst Point\nGiven Offset";

   sOffset = ""; sOffsetType = ""; sZipLength = ""; sZipType = "" ;    
   nMedial = 0 ;
   while (Loop) {		
            nButX = 75; nButY = 80;
            hDlg = Open ("dialog", "General Zipper", "font Helv 8", 20,20, 242, 97);

	AddControl ( hDlg, "pushbutton",  nButX, nButY, 35, 14, "Cancel", "%cancel", "");
	AddControl ( hDlg, "pushbutton",  nButX + 48 , nButY, 35,14, "OK", "%ok", "");

	AddControl ( hDlg, "ltext", 5, 12, 55, 14 , "Draw using:", "string", "");
	AddControl ( hDlg, "dropdownlist", 60, 10, 65, 40 , sDlgLengthList , "string", "sZipType");
	AddControl ( hDlg, "ltext", 135, 12, 35, 14 , "Length:", "string", "");
	AddControl ( hDlg, "ledit", 170, 10, 25, 14 , "" , "string", "sZipLength");

	AddControl ( hDlg, "ltext", 5, 30, 55, 14 , "Profile Offset:", "string", "");
	AddControl ( hDlg, "dropdownlist", 60, 28, 65, 40 , sDlgProfileOffset  , "string", "sOffsetType");
	AddControl ( hDlg, "ltext", 135, 30, 30, 14 , "Offset:", "string", "");
	AddControl ( hDlg, "ledit", 170, 28, 25, 14 , ""  , "string", "sOffset");

	AddControl ( hDlg, "ltext", 5, 52, 25, 14 , "Distal :", "string", "");
	AddControl ( hDlg, "dropdownlist", 30, 50, 55, 70 , sDlgDistalElasticList , "string", "sElasticDistal");

	AddControl ( hDlg, "ltext", 90, 52, 30, 14 , "Proximal:", "string", "");
 	AddControl ( hDlg, "dropdownlist", 125, 50, 55, 70 , sDlgProximalElasticList , "string", "sElasticProximal");
 	
  	AddControl ( hDlg, "checkbox", 185, 50, 65, 15 , "Medial Zipper", "number", "nMedial");

   	Ok = Display("dialog", hDlg, "%center");
  	Close("dialog", hDlg);

	if (Ok == %cancel ) Exit (%ok, "User Cancel!") ;	

   	TwoPoint = %false ;
	PointLength = %false ;
	FirstPoint = %false ;
	StandardOffset = %false ;
	GivenOffset = %false ;

	if (StringCompare("Two Points", sZipType)) TwoPoint = %true;
	if (StringCompare("Point & Length", sZipType)) PointLength = %true;
	if (StringCompare("1-1/8\"", sOffsetType)) StandardOffset = %true;
	if (StringCompare("First Point", sOffsetType)) FirstPoint = %true;
	if (StringCompare("Given Offset", sOffsetType)) GivenOffset = %true;
	    
	// Check that length data is valid
	if (TwoPoint)
		Loop = %false ;
	else {
		nZipLength = Value("scalar", sZipLength);
		if ( nZipLength == 0 ) {
			Display ("beep", "Invalid given length!") ;
			Loop = %true ;
			}
   		else
			Loop = %false ;
		}
	// check that offset data is valid
	if (!Loop) {
		if (StandardOffset || FirstPoint){
			Loop = %false ;
			nZipOff = 1.125 ;
			}
		else {
			nZipOff = Value("scalar", sOffset);
			if ( nZipOff == 0) {
				Display ("beep", "Invalid Offset Given!") ;
				Loop = %true ;
				}
   			else
				Loop = %false ;
			}
		}
	
    }   

 Display ("cursor", "wait", "Searching for more data");

// Get profile, identify Leg and FileNo
//
   if (StandardOffset || GivenOffset ) {
	if(!hCurve = GetEntity ("Select a Leg Profile", "curve")) Exit (%ok, "User cancelled") ;

	Display ("cursor", "wait", "Searching for more data");
   	GetDBData (hCurve , "ID", &sProfileID);
   	nStringLength = StringLength(sProfileID);
   	if (StringCompare( "LeftLegCurve", StringMiddle(sProfileID, nStringLength - 11, 12))){
		sStyleID = StringMiddle(sProfileID, 1, nStringLength - 12);
		sLeg = "Left";
		}
   	if (StringCompare( "RightLegCurve", StringMiddle(sProfileID, nStringLength - 12, 13))) {
		sStyleID = StringMiddle(sProfileID, 1, nStringLength - 13);
		sLeg = "Right";
		}

	// Check if this is a valid Leg Curve exit if not
   	 if (StringLength (sLeg) == 0) Exit (%cancel, "A Leg Profile was not selected\n") ;
// Get Marker data
   sOtemplateID = sStyleID + sLeg + "Origin";
   sAnkleID        =  sStyleID + sLeg + "Ankle";
   sTmp = "type = 'Marker' AND (DB ID ='"  + sAnkleID  + "' OR DB ID ='" + sOtemplateID + "')";
   nMarkersFound = 0 ;
   FootLess = %true ;
   OldDrawing = %false ;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
		GetGeometry(hEnt, &sTmp1, &xyPt1, &xyTmp, &aAngle) ;
	                if ( xyPt1.x != 0 || xyPt1.y != 0) {
			if (StringCompare(sTmp, sOtemplateID)) {
				nMarkersFound = nMarkersFound + 1 ;
				GetGeometry(hEnt, &sTmp, &xyOtemplate , &xyTmp, &aAngle) ;
				GetDBData(hEnt, "units", &sUnits);
				GetDBData(hEnt, "Data", &sTmp) ; // If Data does not exist sTmp will be unchanged
				if ( StringCompare(StringMiddle(sTmp,1,1),"X")) {
					OldDrawing = %true ;
					break;
					}
	  			if ( !ScanLine(sTmp, "blank", &nType, &nn)) Exit(%cancel,"Can't get data from Origin Marker!\n");				
				}
			if (StringCompare(sTmp, sAnkleID)) {
				FootLess = %false ;
				GetGeometry(hEnt, &sTmp, &xyAnkle , &xyTmp, &aAngle) ;
				}
			}
		}
   	}
   Close("selection", hChan); 

// Check if that the markers have been found, otherwise exit
//
    if (nMarkersFound < 1) Exit (%cancel, "Missing marker for selected style, data not found!\n") ;
    if (nMarkersFound > 1) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;

// Get Smallest circumference by looping through the vetex of the curve 
// For legs with a foot then use ankle
   if (FootLess) { 
   	xySmallestCir.y = 10000 ; // Impossibly large value used as an initial value  
   	nn = 1;
   	while (nn < GetVertexCount(hCurve)) {
        		GetVertex(hCurve, nn, &xyTmp) ;
		if ( xyTmp.y < xySmallestCir.y)
			xySmallestCir.y = xyTmp.y ;
		else
			break ; 
		nn = nn + 1;
		}
	}
    else 
    	xySmallestCir.y = xyAnkle.y ;
    	
// Get Smallest circumference if old drawing
//
   if (OldDrawing) {  
   	Display ("message", "error", "This Markers for this leg do not contain enough data"
   				+ "to establish the Smallest Circumference!\n Select manually.") ;
   	if (!GetUser("xy", "Select smallest leg circumference", &xySmallestCir))Exit (%cancel, "Smallest Circumference not selected\n") ;
	}
   }
// Establish allowance for zippers
//
   nElasticProximal = 0.75 ;
   if (StringCompare(StringMiddle(sElasticProximal,1,1), "N")) nElasticProximal = 0 ;
   if (StringCompare(StringMiddle(sElasticProximal,1,3), "3/8")) nElasticProximal = 0.375 ;
   if (StringCompare(StringMiddle(sElasticProximal,1,1), "1")) nElasticProximal = 1.5 ;

   nElasticDistal = 0.75 ;
   if (StringCompare(StringMiddle(sElasticDistal,1,1), "N")) nElasticDistal = 0 ;
   if (StringCompare(StringMiddle(sElasticDistal,1,3), "3/8")) nElasticDistal = 0.375 ;
   if (StringCompare(StringMiddle(sElasticDistal,1,1), "1")) nElasticDistal = 1.5 ;

   nElastic = nElasticDistal + nElasticProximal ;

// Draw on layer notes
//
   hLayer = Table("find", "layer", "Notes") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Notes in Green", "current", Table("find", "color", "Green"));

// Get X values
//
   if (TwoPoint){
   	if (!GetUser("xy", "Select Start of Zipper", &xyHorizStart)) Exit(%ok, "Start not selected") ;
	if (!GetUser("xy", "Select End of Zipper", &xyHorizEnd)) Exit(%ok, "End not selected") ;
	if (xyHorizEnd.x < xyHorizStart.x){
		xyTmp = xyHorizEnd;
		xyHorizEnd = xyHorizStart ;
		xyHorizStart = xyTmp ;
		}
	}

   if (PointLength) 	
  	if (!GetUser("xy", "Select Start of Zipper", &xyHorizStart)) Exit(%ok, "Start not selected") ;

  
// Y values
//
   if (FirstPoint) 
   	xyHorizEnd.y = xyHorizStart.y;
   else {
   	xyHorizStart.y = xySmallestCir.y - nZipOff ;
    	xyHorizEnd.y = xyHorizStart.y;
    	}
 
// Calculate Zipper length
//      	
   if (nZipLength > 0 )
	xyHorizEnd.x = xyHorizStart.x + ((nZipLength / 1.2 ) - nElastic) ;
   else	{
   	nDrawnLength = Calc("length",xyHorizStart, xyHorizEnd) ;
	nZipLength = (nDrawnLength + nElastic ) * 1.2 ;
	}
	
  
   if (nMedial == 1 )
   	sTmp = Format("length",nZipLength) +  " MEDIAL ZIPPER" ;
   else
      	sTmp =  Format("length",nZipLength) +  " LATERAL ZIPPER" ;

   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB") ;   	     
   Symbol( "find", "TextAsSymbol") ;
   hTextEnt = AddEntity ("symbol", "TextAsSymbol", xyHorizStart.x + (xyHorizEnd.x- xyHorizStart.x)/2, xyHorizStart.y + .25, 1, 1, 0) ;  

// Create ID string
   sZipperID =  sStyleID + MakeString("scalar", UID("get", hTextEnt )) ;
   SetDBData(hTextEnt, "ID", sZipperID );
   SetDBData(hTextEnt, "Zipper", "1" );
   SetDBData(hTextEnt, "Data", sTmp ); 

// Draw zip line
   hEnt =  AddEntity("line",xyHorizStart ,xyHorizEnd) ;
   PR_SetZipperID () ;
   	
// Add label and arrows
   hEnt =  AddEntity("marker","closed arrow", xyHorizStart , 0.5 ,0.125, 0 ) ;	
   PR_SetZipperID () ;
   hEnt =  AddEntity("marker","closed arrow",xyHorizEnd , 0.5 ,0.125, 180 ) ;	
   PR_SetZipperID () ;

// Reset and exit
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Zipper drawing Complete");
