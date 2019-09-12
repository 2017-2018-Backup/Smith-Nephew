// File Name:   BDYBUTTN.D
// Purpose:     Create User-Defined Button Box for
//             	       use with Jobst System - Body Brief/Suit with/without sleeves
//                     BODY features.
//
// Version:	1.00
// Date:		7/5/96
// Author:		Paul O'Rawe
//---------------------------------------------------------------------------------------------
// REVISIONS:   
// Date By      Action
//---------------------------------------------------------------------------------------------
//
//
// initialize
STRING 	sPathJOBST;
sPathJOBST = "@" + GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

UserButtons("init", "all", "BodySuit");
// run macros
UserButtons("add", "macro", " -- BODY -- ", "");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Body", "Execute (\"application\", \"jobstart bodysuit\", \"normal\");");
UserButtons("add", "macro", "Sel. Body", "Execute (\"application\", \"jobstart bodysuit S\", \"normal\");");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Left Leg",  "Execute (\"application\", \"jobstart bodyleginput L\", \"normal\");" );
UserButtons("add", "macro", "Right Leg",  "Execute (\"application\", \"jobstart bodyleginput R\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Draw Body", "Execute (\"application\", \"jobstart bodydraw\", \"normal\");");
UserButtons("add", "macro", "Mirr. Profile",sPathJOBST + "\\BODY\\MIRPROF.D");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Left Slv.", "Execute (\"application\", \"jobstart bodyarminput L\", \"normal\");");
UserButtons("add", "macro", "Right Slv.", "Execute (\"application\", \"jobstart bodyarminput R\", \"normal\");");
UserButtons("add", "macro", "Sel. Slv.", "Execute (\"application\", \"jobstart bodyarminput S\", \"normal\");");
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Edit Slv.","Execute (\"application\", \"jobstart armedit\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Tools", sPathJOBST + "\\BUTTONS\\BDYTOOLS;");
UserButtons("add", "macro", "Main Menu", sPathJOBST + "\\BUTTONS\\JOBST;");
// [EndOfMacro]
