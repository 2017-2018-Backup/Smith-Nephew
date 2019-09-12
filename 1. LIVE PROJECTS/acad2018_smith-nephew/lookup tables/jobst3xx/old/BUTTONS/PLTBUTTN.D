// File Name:	PLTBUTTN.D
// Purpose:	Create User-Defined Button Box for
//		To switch non esential layers on and off
//		set up page and printer
//		to plot
//
// Version:	1.01
// Date:		14.Nov.93
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

UserButtons("init", "all", "PLOT");

// run macros
UserButtons("add", "macro", " -- Plot -- ", "");
UserButtons("add", "macro", "Layers Off", sPathJOBST  + "\\PLOT\\plotloff;");
UserButtons("add", "macro", "Layers On", sPathJOBST  + "\\PLOT\\plotlon;");
UserButtons("add", "macro", "Title Off", sPathJOBST  + "\\PLOT\\titleoff;");
UserButtons("add", "macro", "Title On", sPathJOBST  + "\\PLOT\\titleon;");
UserButtons("add", "macro", "","" );
UserButtons("add", "menu", "Page Setup", "FilePageSetup");
UserButtons("add", "macro", "","" );
UserButtons("add", "menu", "Plot", "FilePrint");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Main Menu",sPathJOBST  + "\\BUTTONS\\JOBST;");
// [EndOfMacro]
