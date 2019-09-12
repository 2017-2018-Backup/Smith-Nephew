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
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Zip-Palmer", sPathJOBST + "\\GLOVECAD\\ZIPPALM;");
UserButtons("add", "macro", "Zip-Dorsal", sPathJOBST + "\\GLOVECAD\\ZIPDORSL;");
UserButtons("add", "macro", "Zip-L.F.S", sPathJOBST + "\\GLOVECAD\\ZIPULNAR;");
UserButtons("add", "macro", "Zip LFS-Elbw", sPathJOBST + "\\GLOVECAD\\ZIPLFS;");
UserButtons("add", "macro", "Zip-Thumb", sPathJOBST + "\\GLOVECAD\\ZIPOUT;");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Select Zip", sPathJOBST + "\\TOOLS\\ZIPSELCT;");
UserButtons("add", "macro", "Edit Zip", sPathJOBST + "\\TOOLS\\ZIPEDIT;");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "CAD Glove ", sPathJOBST + "\\BUTTONS\\GLBUTCAD; ");
// [EndOfMacro]
