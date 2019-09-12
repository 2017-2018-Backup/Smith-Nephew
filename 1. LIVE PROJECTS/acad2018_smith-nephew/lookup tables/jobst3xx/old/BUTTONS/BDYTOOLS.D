// File Name:	BDYTOOLS.D
// Purpose:	Create User-Defined Button Box 
//		BodySuit  Tool options.
//
// Version:	1.0
// Date:		20.Jul,95
// Author:		Paul O'Rawe
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
//---------------------------------------------------------------------------------------------
//
// 
// initialize
STRING sPathJOBST;
sPathJOBST = "@" + GetProfileString("JOBST", "PathJOBST", "C:\\JOBST", "DRAFIX.INI") ;

UserButtons("init", "all", "BODY - Tools");

// run macros
UserButtons("add", "macro", " -- TOOLS -- ", " ");
UserButtons("add", "macro", "Arrow", sPathJOBST + "\\TOOLS\\TL_ARROW;");
UserButtons("add", "macro", "Stamps", sPathJOBST + "\\STAMPS\\STMPSTAM;" );
UserButtons("add", "macro", "", "");
UserButtons("add", "macro", "CutOut@Axilla", sPathJOBST + "\\BODY\\EDITPROF;");
UserButtons("add", "macro", "Thumb Hole", sPathJOBST + "\\ARM\\ARMTHUMB;" );
UserButtons("add", "macro", "", "");
UserButtons("add", "macro", "Zip-Latrl", sPathJOBST + "\\TOOLS\\TL_ZLARM; " );
UserButtons("add", "macro", "Zip-Medl", sPathJOBST + "\\TOOLS\\TL_ZMARM; " );
UserButtons("add", "macro", "Zip-Plmr", sPathJOBST + "\\TOOLS\\TL_ZPARM;" );
UserButtons("add", "macro", "Zip-Drsl", sPathJOBST + "\\TOOLS\\TL_ZDARM;" );
UserButtons("add", "macro", "", "");
UserButtons("add", "macro", "Zip-Body", sPathJOBST + "\\BODY\\BDYZIP; " );
UserButtons("add", "macro", "", "");
UserButtons("add", "macro", "Select Zip", sPathJOBST + "\\TOOLS\\ZIPSELCT; " );
UserButtons("add", "macro", "Edit Zip", sPathJOBST + "\\TOOLS\\ZIPEDIT; " );
UserButtons("add", "macro", "BodySuit Menu", sPathJOBST + "\\BUTTONS\\BDYBUTTN;");
// [EndOfMacro]
