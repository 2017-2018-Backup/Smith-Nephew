// File Name:	GLTLSMAN.D
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
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Thumb", sPathJOBST + "\\GLOVEMAN\\GLVTHUMB;");
UserButtons("add", "macro", "Insert", sPathJOBST + "\\GLOVEMAN\\GLVINSRT;");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Zip-Palmer", sPathJOBST + "\\GLOVEMAN\\ZIPPALM;");
UserButtons("add", "macro", "Zip-Dorsal", sPathJOBST + "\\GLOVEMAN\\ZIPDORSL;");
UserButtons("add", "macro", "Zip-L.F.S", sPathJOBST + "\\GLOVEMAN\\ZIPULNAR;");
UserButtons("add", "macro", "Zip-Thumb", sPathJOBST + "\\GLOVEMAN\\ZIPOUT;");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Select Zip", sPathJOBST + "\\GLOVEMAN\\ZIPSELCT;");
UserButtons("add", "macro", "Edit Zip", sPathJOBST + "\\TOOLS\\ZIPEDIT;");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Glove Menu", sPathJOBST + "\\BUTTONS\\GLBUTMAN; ");
// [EndOfMacro]
