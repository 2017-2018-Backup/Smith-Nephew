// File Name:	ARMBUTTN.D
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
sPathJOBST = "@" + GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI");

UserButtons("init", "all", "ARMS");
// run macros
UserButtons("add", "macro", " -- ARMS -- ", "");
UserButtons("add", "macro", "Left",  "Execute (\"application\", \"jobstart arminput L\", \"normal\");" );
UserButtons("add", "macro", "Right", "Execute (\"application\", \"jobstart arminput R\", \"normal\");" );
UserButtons("add", "macro", "Sel Multi", "Execute (\"application\", \"jobstart arminput S\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Edit", "Execute (\"application\", \"jobstart armedit\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Tools",sPathJOBST + "\\BUTTONS\\ARMTOOLS;" );
UserButtons("add", "macro", "Main Menu",sPathJOBST + "\\BUTTONS\\jobst;");

// [EndOfMacro]
