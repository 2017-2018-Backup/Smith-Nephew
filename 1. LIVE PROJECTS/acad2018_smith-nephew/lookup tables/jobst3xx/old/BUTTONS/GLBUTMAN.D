// File Name:	GLBUTMAN.D
// Purpose:	Create User-Defined Button Box 
//		Manual Glove options.
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
sPathJOBST = "@" + GetProfileString("JOBST", "PathJOBST", "C:\\JOBST", "DRAFIX.INI") ;

UserButtons("init", "all", "Gloves");
// run macros
UserButtons("add", "macro", "- Man Glv. -", "");
UserButtons("add", "macro", "MAN - Left",  "Execute (\"application\", \"jobstart manglove L\", \"normal\");" );
UserButtons("add", "macro", "MAN - Right", "Execute (\"application\", \"jobstart manglove R\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Sel Multi", "Execute (\"application\", \"jobstart manglove S\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "WEB - Left",  "Execute (\"application\", \"jobstart webspacr L\", \"normal\");" );
UserButtons("add", "macro", "WEB - Right",  "Execute (\"application\", \"jobstart webspacr R\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Edit", "Execute (\"application\", \"jobstart manglove E\", \"normal\");" );
UserButtons("add", "macro", "","");
UserButtons("add", "macro", "Tools", sPathJOBST + "\\BUTTONS\\GLTLSMAN;");
UserButtons("add", "macro", "Main Menu", sPathJOBST + "\\BUTTONS\\jobst;");
// [EndOfMacro]
