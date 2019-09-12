// File Name:	JOBSTOLD.D
// Purpose:	Swaps DRAFIX back to the old JOBST CAD System
//		v 1.00
//
// Version:	1.0
// Date:		9.Aug.95
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
//---------------------------------------------------------------------------------------------
//
//
// run macros

// initialize
   STRING sPathWINDOWS,
   	 sPathJOBST;
   	
   NUMBER Ok;
   	
   sPathWINDOWS = GetProfileString("JOBST", "PathWINDOWS", "C:\\WINDOWS", "DRAFIX.INI") ;

   sPathJOBST = "C:\\JOBST" ;

// Warn that the user must restart drafix
//
   Ok =  Display ("message", "okquestion",
   			   "Changing back to original JOBST CAD System\n"
   			+ "\nDRAFIX will shut down and will need to be restarted\n"
			+ "Use CANCEL to abort operation");

   if (Ok == %cancel)Exit(%ok, "User Cancelled operation");
   
// Write to the DRAFIX.INI the required changes
 
   SetProfileString("Path", "PathSymbol", sPathJOBST, "DRAFIX.INI") ;
   SetProfileString("Path", "PathMacro", sPathJOBST, "DRAFIX.INI") ;

   SetProfileString("Startup", "DefaultSLB", sPathJOBST + "\\JOBST.SLB", "DRAFIX.INI") ;

   SetProfileString("Macro", "AutoExec", sPathJOBST + "\\JOBSTEXC.D", "DRAFIX.INI") ;
   SetProfileString("Macro", "AutoNew", sPathJOBST + "\\JOBSTNEW.D", "DRAFIX.INI") ;
   SetProfileString("Macro", "AutoClose", "", "DRAFIX.INI") ;
   SetProfileString("Macro", "AutoOpen", "", "DRAFIX.INI") ;
   SetProfileString("Macro", "OutputLog", "", "DRAFIX.INI") ;
   
   SetProfileString("JOBST", "PathJOBST",sPathJOBST, "DRAFIX.INI") ;
   SetProfileString("JOBST", "TXFMacro", "", "DRAFIX.INI") ;

   SetProfileString("Window", "Maximized", "1", "DRAFIX.INI") ;
   SetProfileString("Window", "DocMaximized", "1", "DRAFIX.INI") ;

// Restart the jobst system
// Start Button menu
   Execute("menu", "FileExit") ;