// File Name:	HNBUTTON.D
// Purpose:	Create User-Defined Button Box 
//		Head and Neck  options.
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

UserButtons("init", "all", "Head & Neck");

// run macros
UserButtons("add", "macro", "Head+Neck", "");
UserButtons("add", "macro", "HeadNeck", sPathJOBST + "\\HEADNECK\\HEADNECK;");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Sel. Multi.", sPathJOBST + "\\HEADNECK\\HEADSEL;");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Main Menu",sPathJOBST + "\\BUTTONS\\jobst;");
// [EndOfMacro]
