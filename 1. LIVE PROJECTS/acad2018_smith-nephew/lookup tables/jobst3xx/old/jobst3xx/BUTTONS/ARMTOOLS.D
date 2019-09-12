// File Name:	ARMTOOLS.D
// Purpose:	Create User-Defined Button Box 
//		Arm  Tool options.
//
// Version:	1.1
// Date:		8.Feb,95
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

UserButtons("init", "all");

// run macros
UserButtons("add", "macro", " -- TOOLS -- ", " ");
UserButtons("add", "macro", "Arrow", sPathJOBST + "\\TOOLS\\TL_ARROW; ");
UserButtons("add", "macro", "Stamps", "Execute (\"application\", \"jobstart stamp am\", \"normal\");" );
UserButtons("add", "macro", "Thumb Hole", sPathJOBST + "\\ARM\\ARMTHUMB;" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Zip-Latrl", sPathJOBST + "\\TOOLS\\TL_ZLARM; " );
UserButtons("add", "macro", "Zip-Medl", sPathJOBST + "\\TOOLS\\TL_ZMARM; " );
UserButtons("add", "macro", "Zip-Plmr", sPathJOBST + "\\TOOLS\\TL_ZPARM;" );
UserButtons("add", "macro", "Zip-Drsl", sPathJOBST + "\\TOOLS\\TL_ZDARM;" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Zip-Gaunt", sPathJOBST + "\\TOOLS\\GAUNTZIP;" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Arm Menu", sPathJOBST + "\\BUTTONS\\armbuttn; ");
// [EndOfMacro]
