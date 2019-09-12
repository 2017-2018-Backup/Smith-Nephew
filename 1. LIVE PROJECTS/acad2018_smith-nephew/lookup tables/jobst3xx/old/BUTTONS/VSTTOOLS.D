// File Name:	VSTTOOLS.D
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

UserButtons("init", "all", "VEST - Tools");

// run macros
UserButtons("add", "macro", " -- TOOLS -- ", " ");
UserButtons("add", "macro", "Arrow", sPathJOBST + "\\TOOLS\\TL_ARROW;");
UserButtons("add", "macro", "Stamps", sPathJOBST + "\\STAMPS\\STMPSTAM;" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Thumb Hole", sPathJOBST + "\\ARM\\ARMTHUMB;" );
UserButtons("add", "macro", "Reglr. Axilla", sPathJOBST + "\\RAGLAN\\FIXREGLR" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Zip-Latrl", sPathJOBST + "\\TOOLS\\TL_ZLARM; " );
UserButtons("add", "macro", "Zip-Medl", sPathJOBST + "\\TOOLS\\TL_ZMARM; " );
UserButtons("add", "macro", "Zip-Plmr", sPathJOBST + "\\TOOLS\\TL_ZPARM;" );
UserButtons("add", "macro", "Zip-Drsl", sPathJOBST + "\\TOOLS\\TL_ZDARM;" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Zip-Gaunt", sPathJOBST + "\\TOOLS\\GAUNTZIP;" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Vest to Slv.",  sPathJOBST + "\\VEST\\BODTOSLV;");
UserButtons("add", "macro", "Draw Point", sPathJOBST + "\\VEST\\BODAXILA;");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Scoop Vest", sPathJOBST + "\\VEST\\SCPVSTIN;");
UserButtons("add", "macro", "Scoop Slv.", sPathJOBST + "\\VEST\\SCPSLVIN;");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Vest Menu", sPathJOBST + "\\BUTTONS\\vstbuttn;");
// [EndOfMacro]
