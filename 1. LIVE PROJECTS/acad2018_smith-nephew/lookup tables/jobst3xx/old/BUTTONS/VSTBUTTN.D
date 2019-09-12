// File Name:   VSTBUTTN.D
// Purpose:     Create User-Defined Button Box for
//              use with Jobst System - Vest With
//              sleeves
//              BODY features.
//
// Version:     2.01
// Date:                20.Oct.93
// Author:              Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:   
// Date By      Action
//---------------------------------------------------------------------------------------------
//
//
// initialize
BOOL	OK;
STRING 	sPathJOBST;
sPathJOBST = "@" + GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

UserButtons("init", "all", "VEST");
// run macros
UserButtons("add", "macro", " - VEST - ", " ") ;
UserButtons("add", "macro", "Vest Details",  "Execute (\"application\", \"jobstart vest\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Left Slv.",  "Execute (\"application\", \"jobstart vestarminput L\", \"normal\");" );
UserButtons("add", "macro", "Right Slv.", "Execute (\"application\", \"jobstart vestarminput R\", \"normal\");" );
UserButtons("add", "macro", "Sel Multi", "Execute (\"application\", \"jobstart vestarminput S\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Edit Slv.","Execute (\"application\", \"jobstart armedit\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Torso Band","Execute (\"application\", \"jobstart torsoband\", \"normal\");" );
//UserButtons("add", "macro", "Scoop Neck","Execute (\"application\", \"jobstart scoopneck\", \"normal\");" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "","" );
UserButtons("add", "macro", "Tools", sPathJOBST + "\\BUTTONS\\VSTTOOLS;");
UserButtons("add", "macro", "Main Menu", sPathJOBST + "\\BUTTONS\\JOBST;");
// [EndOfMacro]
