// File:		GLVINSRT.D
// Purpose:	Insert an Insert symbol.
// Version:	1.1
// Date:		14.May.96
// Author		G.George
//		C-Gem Software
// Modifications
// Date		By	Reason

XY
	xyOrigin,
	xyPoint;
ANGLE
	aAngle;

BOOL
	PointGiven ;
	
STRING	sPathJOBST, sText;
	
   UserSelection ("clear") ;
   UserSelection ("update") ;
   
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB") ;   	     

// Set to Layer Notes
//
   Execute("menu","SetLayer",Table("find","Layer","Notes"));

// Insert symbol (Loop until user presses Right Mouse button)
// Get point from user

   GetUser("xy","Origin for insert",&xyOrigin) ;
   Symbol( "find", "GloveInsertU") ;
   hEnt = AddEntity ("symbol", "GloveInsertU", xyOrigin) ;  
   if (GetUser("string", "Text to attach to insert Symbol", &sText)) {
   	SetData("TextVertJust", 8) ;
   	SetData("TextHorzJust", 2) ;   		
   	SetData("TextHeight", 0.125) ;   		
   	AddEntity ("text", sText, xyOrigin.x ,xyOrigin.y - 0.25) ;  
  	}
 
// Reset to Layer 1
//
   Execute("menu","SetLayer",Table("find","Layer","1"));




	