// File Name:	LGBUTTON.D
// Purpose:	Create User-Defined Button Box 
//		Arm  options.
//
// Version:	1.01
// Date:		31.Oct.94
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
//---------------------------------------------------------------------------------------------
//
// 
// initialize
STRING sPathJOBST;
sPathJOBST = "@" + GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

UserButtons("init", "all", "LEGS");

// run macros
UserButtons("add", "macro", " -- LEGS -- ", "");
UserButtons("add", "macro", "Left", "Execute (\"application\", \"jobstart leginput L\", \"normal\");" );
UserButtons("add", "macro", "Right",  "Execute (\"application\", \"jobstart leginput R\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Leg Tools", sPathJOBST  + "\\BUTTONS\\LGTOOLS; " );
UserButtons("add", "macro", "Main Menu", sPathJOBST  + "\\BUTTONS\\JOBST; ");
// [EndOfMacro]
