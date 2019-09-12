// File Name:	WHBUTN1.D
// Purpose:	Create User-Defined Button Box for
//		Waist Height with two legs.
//
// Version:	1.01
// Date:		11.Jan.94
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

UserButtons("init", "all", "WAIST");

// run macros
UserButtons("add", "macro", " -- WAIST -- ", "");
UserButtons("add", "macro", "",  "");
UserButtons("add", "macro", "Body",  "Execute (\"application\", \"jobstart waistbody\", \"normal\");" );
UserButtons("add", "macro", "",  "");
UserButtons("add", "macro", "Left Leg",  "Execute (\"application\", \"jobstart waistleg L\", \"normal\");" );
UserButtons("add", "macro", "Right Leg",  "Execute (\"application\", \"jobstart waistleg R\", \"normal\");" );
UserButtons("add", "macro", "Figure",  "Execute (\"application\", \"jobstart waistfigure\", \"normal\");" );
UserButtons("add", "macro", "",  "");
UserButtons("add", "macro", "Cut-Out", sPathJOBST + "\\WAIST\\WH_DRAW WHCUT ; ");
UserButtons("add", "macro", "1st Leg", sPathJOBST + "\\WAIST\\WH_DRAW WH1LG ; ");
UserButtons("add", "macro", "2nd Leg", sPathJOBST + "\\WAIST\\WH_DRAW WH2LG ; ");
UserButtons("add", "macro", "Label", sPathJOBST + "\\WAIST\\wh_Labl ; ");
UserButtons("add", "macro", "Chap", sPathJOBST + "\\WAIST\\WH_CHAP ; ");
UserButtons("add", "macro", "",  "");
UserButtons("add", "macro", "Edit","Execute (\"application\", \"jobstart waistlegedit\", \"normal\");" );
UserButtons("add", "macro", "WH Tools", sPathJOBST + "\\BUTTONS\\WHTOOLS ; " );
UserButtons("add", "macro", "Main Menu", sPathJOBST + "\\BUTTONS\\jobst ; ");
// [EndOfMacro]
