// File:		TL_ZDARM.D
// Purpose:	Insert Dorsal Zipper on Arm.
// Version:	1.1
// Date:		8.Feb.95
// Author		G.Dunne
//		Jobst Ireland.

XY
	xyZipperStart,
	xyZipperEnd,
	xyTextPos;

NUMBER
	nActualLen,
	nFiguredLen;
	
STRING
	sFiguredLen;



// Get Start point of zipper 
// Get point from user
//
   Execute("Menu", "SnapMidpoint") ;
   GetUser("xy","Select Mid point of Distal End",&xyZipperStart);

// Get End Point of Zipper
//
   Execute("Menu", "SnapOff") ;
   GetUser("xy","Select End Point of Zipper",&xyZipperEnd);

// Check that User has given Start point at Distal end.
// Swap Start and End points, if given Start point is at Proximal end.
//
   if (xyZipperStart.x > xyZipperEnd.x) {
	Display("message", "error", "WARNING!\nStart Zipper at Distal End\n\n");
	Exit(%ok, "");
	}

//   Calculate  Figured Length
//	
   nActualLen = xyZipperEnd.x - xyZipperStart.x ;
   nFiguredLen=nActualLen/0.92;
   sFiguredLen=Format("length",nFiguredLen)+" Dorsal Zipper";

// Check if Actual length too small to draw
// Beep and exit
//
   nLenTol = 1 ; 
   if (nActualLen < nLenTol) 
	Exit(%ok, "Too Small to Draw!");

// Revise Zipper Start point to make allowance for seam
//
   nSeam =  0.1875/ 2 ;
   xyZipperStart.y = xyZipperStart.y + nSeam;

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

// Insert Line
//
   AddEntity("line", xyZipperStart, xyZipperEnd.x, xyZipperStart.y) ;

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






	