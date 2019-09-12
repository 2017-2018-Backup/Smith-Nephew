// File Name:	JOBSTMNU.D
// Purpose:	Create User-Defined Menu
//		Plotting and Titlebox
//
// Version:	1.02
// Date:		16.Mar.95
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
//---------------------------------------------------------------------------------------------
//
//
// run macros

// initialize
STRING sPathJOBST;
sPathJOBST = "@" + GetProfileString("JOBST", "PathJOBST", "C:\\JOBST", "DRAFIX.INI") ;

SetLastMenu( "JOBST",  
	"Open", "Execute(\"menu\",\"FileOpen\");",
	"Close", sPathJOBST  + "\\CADLINK\\CL_CLOSE; ",
	"Save", sPathJOBST  + "\\CADLINK\\CL_SAVE; ",
	"---------------", "",
	"Title Box", "Execute (\"application\", \"jobstart PDandWOD\", \"normal\");",
	"---------------", "",
	"Layers Off",sPathJOBST  + "\\PLOT\\plotloff;",
	"Layers On",sPathJOBST  + "\\PLOT\\plotlon;",
	"Title Off", sPathJOBST  + "\\PLOT\\titleoff;",
	"Title On", sPathJOBST  + "\\PLOT\\titleon;",
	"Page Setup", "Execute (\"menu\",\"FilePageSetup\");",
	"Plot", "Execute(\"menu\",\"FilePrint\");",
	"---------------", "",
	"Special Symbol", sPathJOBST + "\\TOOLS\\TL_WARN; ",
	"MULTI", sPathJOBST + "\\TOOLS\\TL_MULTI.D MULTI; ",
	"TAN", sPathJOBST + "\\TOOLS\\TL_FABRC.D TAN; ",
	"CREAM", sPathJOBST + "\\TOOLS\\TL_FABRC.D CREAM; ",
	"BLACK", sPathJOBST + "\\TOOLS\\TL_FABRC.D BLACK; ",
	"DENIM", sPathJOBST + "\\TOOLS\\TL_FABRC.D DENIM; ",
	"WHITE", sPathJOBST + "\\TOOLS\\TL_FABRC.D WHITE; "
	 );
// [EndOfMacro]
 