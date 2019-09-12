// File:		TL_WARNING.D
// Purpose:	Insert a warning symbol w.r.t a Special Instruction.
// Version:	1.0
// Date:		28.May.97
// Author		G.George
//		C-gem Ltd
// Modifications
// Date		By	Reason

XY
	xyOrigin; 
STRING	 
	sPathJOBST;

// Find JOBST installed directory
// Set path to macros
//
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB") ;   	     
   Symbol( "find", "WarningSpecial") ;
   
// Set to Layer Notes
//
   Execute("menu","SetLayer",Table("find","Layer","Notes"));

// Set insert position (Loop until user presses Right Mouse button)
// Get point from user

   while (GetUser("xy","Select position for warning",&xyOrigin)) {
	// Insert Symbol
   	AddEntity ("symbol", "WarningSpecial", xyOrigin, 1, 1, 0) ;  
   	}

// Reset to Layer 1
//
   Execute("menu","SetLayer",Table("find","Layer","1"));
   Exit (%ok, "TL_WARN, Finished") ;




	