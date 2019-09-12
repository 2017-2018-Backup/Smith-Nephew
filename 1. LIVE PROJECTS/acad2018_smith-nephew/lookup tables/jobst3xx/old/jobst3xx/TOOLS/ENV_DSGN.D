// File Name:	ENV_DSGN.D
// Purpose:	To set up DRAFIX for DESIGN only
//		
// Version:	1.01
// Date:		15.July.93
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
//
STRING sPathJOBST;

//	sPathJOBST = "Y:\\JOBST";
	sPathJOBST = "E:\\JOBST\\JOBST200";
   	SetProfileString("Path", "PathSymbol", sPathJOBST, "DRAFIX.INI");
   	SetProfileString("Path", "PathMacro", sPathJOBST, "DRAFIX.INI");
   	SetProfileString("Startup", "DefaultSLB", sPathJOBST + "\\JOBST.SLB", "DRAFIX.INI") ;
   	SetProfileString("Macro", "AutoExec", sPathJOBST + "\\JOBSTEXC.D", "DRAFIX.INI") ;
   	SetProfileString("Macro", "AutoNew", "", "DRAFIX.INI") ;
   	SetProfileString("Macro", "AutoClose", "", "DRAFIX.INI") ;
   	SetProfileString("Macro", "AutoOpen", "", "DRAFIX.INI") ;
      	SetProfileString("JOBST", "PathJOBST",sPathJOBST, "DRAFIX.INI") ;	

  SetData("PathMacro", GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") + "\\BUTTONS" );
	 
@JOBSTMNU;
@JOBST;

//Exit
