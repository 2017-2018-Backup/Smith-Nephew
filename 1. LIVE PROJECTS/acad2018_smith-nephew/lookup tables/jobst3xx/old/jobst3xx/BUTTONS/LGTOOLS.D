// File Name:	LGTOOLS.D
// Purpose:	Create User-Defined Button Box 
//		Leg Tool options.
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

UserButtons("init", "all", "LEG - Tools");

// run macros
UserButtons("add", "macro", " - TOOLS - ", "; ");
UserButtons("add", "macro", "Reinfc", sPathJOBST + "\\WAIST\\WHFTREFC; ");
UserButtons("add", "macro", "Stamps", "Execute (\"application\", \"jobstart stamp lg\", \"normal\");" );
UserButtons("add", "macro", "Stump", sPathJOBST + "\\LEG\\LG_STUMP; ");
UserButtons("add", "macro", "Knee ctr", sPathJOBST + "\\LEG\\LGCNKNEE; ");
UserButtons("add", "macro", "Heel ctr", sPathJOBST + "\\LEG\\LGCNHEEL; ");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Zip-Ank", sPathJOBST + "\\LEG\\LGZIPANK; " );
UserButtons("add", "macro", "Zip-Proxml", sPathJOBST + "\\LEG\\LGZIPPRX; " );
UserButtons("add", "macro", "Zip-Distal", sPathJOBST + "\\LEG\\LGZIPDST; " );
UserButtons("add", "macro", "Zip", sPathJOBST + "\\LEG\\WHLG_ZIP; " );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Leg Menu", sPathJOBST + "\\BUTTONS\\LGBUTTON; ");
// [EndOfMacro]
