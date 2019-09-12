// File:		TL_FABRC.D <TEXT>
// Purpose:	Insert given text in red
//
// Version:	1.0
// Date:	29.Sept.97
// Author	G.George
//		C-gem Ltd
// Modifications
// Date		By	Reason

XY
	xyOrigin; 

STRING	 
	 sText;
	 
HANDLE
	hEnt;

// Reset 
   UserSelection ("clear");
   UserSelection("update") ;
   Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));

// Set to Layer Notes and Set color to red
//
   Execute("menu","SetLayer",Table("find","Layer","Notes"));
   Execute ( "menu", "SetColor", Table( "find",  "color", "red"));
   
// Set text defaults
//   
   SetData ("TextVertJust", 32);
   SetData ("TextHorzJust", 1);
   SetData ("TextHeight", .25);
   SetData ("TextAngle", 0);
  
    sText = GetArgument (1) ;
   	
// Set insert position (Loop until user presses Right Mouse button)
// Get point from user

   while (GetUser("xy","Select position for text",&xyOrigin)) {
	// Insert Symbol
   	hEnt = AddEntity ("text",sText, xyOrigin) ;  
   	}

// Reset to Layer 1 etc...
//
   Execute ( "menu", "SetColor", Table( "find",  "color","bylayer"));
   Execute("menu","SetLayer",Table("find","Layer","1"));
   Exit (%ok, "TL_FABRC, Finished") ;




	