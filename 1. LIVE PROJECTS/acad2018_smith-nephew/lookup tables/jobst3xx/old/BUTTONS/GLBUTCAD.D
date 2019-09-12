// File Name:	GLBUTCAD.D
// Purpose:	Create User-Defined Button Box 
//			CAD Glove options.
//
// Version:	1.01
// Date:		07.Jul.95
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

UserButtons("init", "all", "Gloves");
// run macros
UserButtons("add", "macro", "- CAD Glv -", "");
UserButtons("add", "macro", "CAD - Left",  "Execute (\"application\", \"jobstart cadglove L\", \"normal\");" );
UserButtons("add", "macro", "CAD - Right", "Execute (\"application\", \"jobstart cadglove R\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Sel Multi", "Execute (\"application\", \"jobstart cadglove S\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "WEB - Left",  "Execute (\"application\", \"jobstart webspacr L\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "WEB - Right",  "Execute (\"application\", \"jobstart webspacr R\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Edit", "Execute (\"application\", \"jobstart cadglove E\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Tools", sPathJOBST + "\\BUTTONS\\GLTLSCAD;");
UserButtons("add", "macro", "Main Menu", sPathJOBST + "\\BUTTONS\\jobst;");
// [EndOfMacro]
