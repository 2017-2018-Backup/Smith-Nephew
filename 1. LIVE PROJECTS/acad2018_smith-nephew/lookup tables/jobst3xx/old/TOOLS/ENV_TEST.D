// File Name:	ENV_TEST.D
// Purpose:	To set up DRAFIX for TESTING only
//	
// Version:	1.01
// Date:		9.Oct.96
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
//
STRING sPathJOBST ;
// sPathJOBST = "C:\\JOBST2xx" ;
sPathJOBST = "E:\\JOBST\\JOBST2xx" ;

// Check that the DRAFIX.INI file allows the specific machine to start
// the test environment
  if (StringCompare(GetProfileString("JOBST", "EnableTestEnvironment", "NO", "DRAFIX.INI"), "YES")) {

	SetProfileString("Path", "PathSymbol", sPathJOBST, "DRAFIX.INI");
	SetProfileString("Path", "PathMacro", sPathJOBST, "DRAFIX.INI");
	SetProfileString("Startup", "DefaultSLB", sPathJOBST + "\\JOBST.SLB", "DRAFIX.INI") ;
	SetProfileString("Macro", "AutoExec", sPathJOBST + "\\JOBSTEXC.D", "DRAFIX.INI") ;
	SetProfileString("Macro", "AutoNew", sPathJOBST + "", "DRAFIX.INI") ;
	SetProfileString("Macro", "AutoClose", "", "DRAFIX.INI") ;
	SetProfileString("Macro", "AutoOpen", "", "DRAFIX.INI") ;
   	SetProfileString("JOBST", "PathJOBST", sPathJOBST, "DRAFIX.INI") ;	
	
	SetData("PathMacro", GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") + "\\BUTTONS" );
	 
	}
  else
	Exit (%cancel, "This workstation does not allow the TEST Environment");

@JOBSTMNU;
@JOBST;
