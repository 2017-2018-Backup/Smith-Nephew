// File Name:	JOBSTEXC.D
// Purpose:	Jobst system startup.
//		See also DRAFIX.INI
//		
// Version:	1.01
// Date:		15.July.93
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
//
// Start Button menu
SetData("PathMacro", GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") + "\\BUTTONS" );
@JOBSTMNU;
@JOBST;
Exit( %ok, "JOBST - Started");

//Exit
