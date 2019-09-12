// File:		TL_ZLARM.D
// Purpose:	Insert Lateral Zipper on Arm.
// Version:	1.1
// Date:		8.Feb.95
// Author		G.Dunne
//		Jobst Ireland.

XY
	xyTmp,
	xyZipperStart,
	xyZipperEnd,
	xyTextPos;

NUMBER
	nHtTol,	// Height Tolerance
	nLenTol,	// Length Tolerance
	nHtDiff,
	nActualLen,
	nFiguredLen;
	
STRING
	sFiguredLen;



// Get Start point of zipper 
// Get point from user
//
   GetUser("xy","Select Start point of Zipper",&xyZipperStart);

// Get End Point of Zipper
//
   GetUser("xy","Select End Point of Zipper",&xyZipperEnd);

// Check that User has given Start point at Distal end.
// Swap Start and End points, if given Start point is at Proximal end.
//
   if (xyZipperStart.x > xyZipperEnd.x) {
	xyTmp = xyZipperStart;
                xyZipperStart = xyZipperEnd;
	xyZipperEnd = xyTmp;
	}
 
// Check that difference in height between ZipperStart and ZipperEnd is less than
// the tolerance
   nHtTol = 0.0625;  // 1/16th of an Inch
   nHtDiff = abs(xyZipperEnd.y - xyZipperStart.y) ;
   if (nHtDiff > nHtTol) {
	Display("message", "error", "WARNING!\nZipper not Horizontal\n\n");
	Exit(%ok, "");
	} 

//   Calculate  Figured Length
//	
   nActualLen=Calc("Length",xyZipperStart,xyZipperEnd);
   nFiguredLen=nActualLen/0.92;
   sFiguredLen=Format("length",nFiguredLen)+" Zipper";

// Check if Actual length too small to draw
// Beep and exit
//
   nLenTol = 1 ; 
   if (nActualLen < nLenTol) 
	Exit(%ok, "Too Small to Draw!");

// Set to Layer Notes
//
   Execute("menu","SetLayer",Table("find","Layer","Notes"));

// Insert Marker
// Markers
//	Type="closed arrow"
//	Length=  0.5inch
//	Width=    0/125inch
//

// Start of Zipper, Marker
//
   AddEntity("marker","closed arrow",xyZipperStart,0.5,0.125,0);

// End of Zipper, Marker
// N.B. Use of Start Point Y value
//
   AddEntity("marker","closed arrow",xyZipperEnd.x, xyZipperStart.y,0.5,0.125,180);

// Calculate Text Insertion Position
//	
   xyTextPos.x=xyZipperStart.x+(nActualLen/2);
   xyTextPos.y=xyZipperStart.y+0.1;

// Set Text Justification and Height
//
   SetData("TextVertJust",32);
   SetData("TextHorzJust",2);
   SetData("TextHeight",0.125);
//
// Insert Text
//
   AddEntity("text",sFiguredLen,xyTextPos);


// Reset to Layer 1
//
   Execute("menu","SetLayer",Table("find","Layer","1"));






	