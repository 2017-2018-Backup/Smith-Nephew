// File Name:	TLSBUTTN.D
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

UserButtons("init", "all", "Tools");

// run macros
UserButtons("add", "macro", " - Tools - ", "");
UserButtons("add", "macro", "Arrow", sPathJOBST + "\\TOOLS\\TL_ARROW; ");
UserButtons("add", "macro", "Stamps", sPathJOBST + "\\STAMPS\\STMPSTRT; " );
UserButtons("add", "macro", "WO Change", sPathJOBST + "\\TOOLS\\WOEDIT; " );
UserButtons("add", "macro", "Special", sPathJOBST + "\\TOOLS\\TL_WARN; " );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Merge", sPathJOBST + "\\CADLINK\\MERGE_IN; " );
UserButtons("add", "macro", "Inquire Sym", sPathJOBST + "\\CADLINK\\MERGEINQ; " );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Main Menu", sPathJOBST + "\\BUTTONS\\JOBST; ");
// [EndOfMacro]
