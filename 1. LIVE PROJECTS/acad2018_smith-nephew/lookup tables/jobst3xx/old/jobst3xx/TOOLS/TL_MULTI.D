// File:		TL_MULTI.D <TEXT>
// Purpose:	Insert given text in blue
//
// Version:	1.0
// Date:	23.Nov.98
// Author	G.Dunne
//		Jobst Ireland Ltd
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

// Set to Layer TemplateLeft and Set color to blue
//
   Execute("menu","SetLayer",Table("find","Layer","TemplateLeft"));
   Execute ( "menu", "SetColor", Table( "find",  "color", "blue"));
   
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




	