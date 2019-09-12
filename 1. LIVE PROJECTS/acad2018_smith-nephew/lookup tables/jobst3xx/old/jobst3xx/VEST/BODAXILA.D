// File Name:	BODAXILA.D
// Purpose:	Draw selected axilla points
//		Used by the operator to draw an axilla
//		point to ensure that the template is 
//		clearly defined
//		N.B. Quick and Dirty
//		
// Version:	1.01
// Date:		28.Jun.94
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------

// Start "HourGlass" cursor
    Display ("cursor", "wait", "Loading Program");
//
   UserSelection ("clear") ;
   UserSelection ("update") ;
   Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));
   Execute ( "menu", "SetColor", Table( "find",  "color", "bylayer"));
    
// Find JOBST installed directory
// Set path to macros
//
   STRING	 sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\VEST" );

// Load variables
//
   @BODYVBLS.D;

// Load Functions and Procedures
//
   @BODYDEF.D;

// Get raglan to get layer only
   if(!hEnt = GetEntity ("Select Vest Raglan Profile", "all"))Exit(%ok,"Vest Raglan Profile not selected");
   GetEntityData ( hEnt, "layer" , &hCurve); 
   GetTableData ( "layer", hCurve, "name", &sTmp);
   sSide = StringMiddle(sTmp, 9, (StringLength(sTmp) - 8));

   if (!StringCompare(sSide,"Left") && !StringCompare(sSide, "Right")) 
	Exit (%cancel, "Select a Right or Left Raglan profile only\n\n") ;

// Get Axilla and  front neck Raglan intersection
// Note - This is point that controls the angle of the raglan curve

   if( !GetUser ("xy", "Axilla Point", &xyHighestAxilla))Exit(%ok, "");
   if( !GetUser ("xy", "Front Neck and Raglan Intersection", &xyFrontNeckOFF))Exit(%ok, "");
   nHighestAxillaFrontNeckRad = Calc ("length", xyHighestAxilla,xyFrontNeckOFF);

// Get the vestcurve angle 
   aCurve = FN_CurveAngle ( xyHighestAxilla, nHighestAxillaFrontNeckRad );
   aVestCurve = Calc ("angle", xyHighestAxilla,xyFrontNeckOFF);

   if (StringCompare (sSide, "Right"))
	PRSetLayerTemplateRight();
   else 
	PRSetLayerTemplateLeft();

// Open Curve for Reading
// Assume that if the tests for Existance/Corruption are passed. Then no need to repeat them

   hCurve = Open ("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly");
   GetLine(hCurve, &sLine);
   ScanLine( sLine, "blank", &nLength, &aAngle);

   aPrevAngle = (aVestCurve - aCurve);      // Rotate curve to correct start angle
   xyCen  = xyHighestAxilla;
   nRadius = 3.5  ;			   // Fudge because I can't be bothered
   xyPt1 = xyHighestAxilla;
   while (GetLine(hCurve, &sLine)){
	aAngle = aAngle + aPrevAngle;
	xyPt2 = CalcXY("relpolar", xyPt1,nLength, aAngle);
	if ( FN_CirLinInt( xyPt1, xyPt2, xyCen, nRadius)) {
			AddEntity("line", xyPt1, xyPt2);
			break;
			}
                AddEntity("line",xyPt1,xyPt2);  // Draw raglan
	xyPt1 = xyPt2;
	aPrevAngle = aAngle;
	ScanLine( sLine, "blank", &nLength, &aAngle);
   } //End GetLine while

   Close ("file", hCurve);

// Finish with "HourGlass" cursor
   Display ("cursor", "nowait", "Axilla point drawing finished");
