// File Name:	FILBUTTN.D
// Purpose:	Create User-Defined Button Box 
//		File options
//
// Version:	1.1
// Date:		8.Feb,95
// Author:		G.Dunne/Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
//---------------------------------------------------------------------------------------------
//
// 
// initialize
STRING sPathJOBST;
sPathJOBST = "@" + GetProfileString("JOBST", "PathJOBST", "C:\\JOBST", "DRAFIX.INI") ;

UserButtons("init", "all", "File");

// run macros
UserButtons("add", "macro", " --- FILE --- ", "");
UserButtons("add", "macro", "","" );
UserButtons("add", "menu", "Open", "FileOpen" );
UserButtons("add", "macro", "Close",  sPathJOBST + "\\CADLINK\\CL_CLOSE; ");
UserButtons("add", "macro", "Save",  sPathJOBST + "\\CADLINK\\CL_SAVE; ");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Title Box","Execute (\"application\", \"jobstart PDandWOD\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "UpDate", sPathJOBST + "\\CADLINK\\CLUPDATE.D; ");
UserButtons("add", "macro", "Error Save", sPathJOBST + "\\CADLINK\\CL_ERROR.D; ");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "TEST Envir.", sPathJOBST + "\\TOOLS\\ENV_TEST.D;");
UserButtons("add", "macro", "DESIGN Envir.", sPathJOBST + "\\TOOLS\\ENV_DSGN.D;");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Main Menu", sPathJOBST + "\\BUTTONS\\JOBST; ");
// [EndOfMacro]
