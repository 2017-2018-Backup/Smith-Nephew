// File:		TL_ARROW.D
// Purpose:	Insert a marker at an angle given by curser.
// Version:	1.2
// Date:		8.Feb.95
// Author		G.Dunne
//		Jobst Ireland.
// Modifications
// Date		By	Reason
// 8.Feb.95	GD/GG	Loop continually

XY
	xyOrigin,
	xyPoint;

ANGLE
	aAngle;

BOOL
	PointGiven ;

// Set to Layer Notes
//
   Execute("menu","SetLayer",Table("find","Layer","Notes"));

// Set Head Of Arrow (Loop until user presses Right Mouse button)
// Get point from user

   while (GetUser("xy","Select Head Of Arrow",&xyOrigin)) {
	// Get End Of Arrow(used to calculate marker angle)
	PointGiven = GetUser("xy","Give End Of Arrow",&xyPoint);	
	if (PointGiven == %false) break;
	// Calculate Angle
   	aAngle=Calc("angle",xyOrigin,xyPoint);
	// Insert Marker
   	AddEntity("marker","closed arrow",xyOrigin,0.5,0.125,aAngle);
   	}

// Reset to Layer 1
//
   Execute("menu","SetLayer",Table("find","Layer","1"));
   Exit (%ok, "TL_ARROW, Finished") ;




	