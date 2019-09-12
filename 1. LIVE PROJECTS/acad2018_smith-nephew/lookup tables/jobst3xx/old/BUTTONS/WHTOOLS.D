// File Name:	WHTLS1.D
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
sPathJOBST = "@" + GetProfileString("JOBST", "PathJOBST", "C:\\JOBST", "DRAFIX.INI") ;

UserButtons("init", "all", "WAIST - Tools");

// run macros
UserButtons("add", "macro", " -- TOOLS -- ", "");
UserButtons("add", "macro", "Reinfc", sPathJOBST + "\\WAIST\\WHFTREFC; ");
UserButtons("add", "macro", "Stamps", "Execute (\"application\", \"jobstart stamp wh\", \"normal\");" );
UserButtons("add", "macro", "Stump",  sPathJOBST + "\\WAIST\\WH_STUMP; ");
UserButtons("add", "macro", "Knee ctr", sPathJOBST + "\\LEG\\LGCNKNEE; ");
UserButtons("add", "macro", "Heel ctr", sPathJOBST + "\\LEG\\LGCNHEEL; ");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Zip-Ank", sPathJOBST + "\\WAIST\\WHZIPANK; " );
UserButtons("add", "macro", "Zip-Lat", sPathJOBST + "\\WAIST\\WHZIPLAT; " );
UserButtons("add", "macro", "Zip-Bod", sPathJOBST + "\\WAIST\\WHZIPBOD; " );
UserButtons("add", "macro", "Zip-Distal", sPathJOBST + "\\WAIST\\WHZIPDST; " );
UserButtons("add", "macro", "Zip",sPathJOBST + "\\WAIST\\WHLG_ZIP; " );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Zip-Chp-Ank", sPathJOBST + "\\WAIST\\WHZIPCHP; " );
UserButtons("add", "macro", "Zip-Chp-Panty", sPathJOBST + "\\WAIST\\WHZCPANT; " );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Select Zip", sPathJOBST + "\\TOOLS\\ZIPSELCT; " );
UserButtons("add", "macro", "Edit Zip", sPathJOBST + "\\TOOLS\\ZIPEDIT; " );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "WH Menu", sPathJOBST + "\\BUTTONS\\WHBUTTON; ");
// [EndOfMacro]
