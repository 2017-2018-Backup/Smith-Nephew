// File Name:	BODL.D
// Purpose: 	Shell for LEFT Sleeve of a body suit
//		Calls BOD_IN.D 
// Version:	1.1
// Date:		10.Sept.97
// By:		Gary George
STRING	sSleeve;
sSleeve = "Left";
SetData("PathMacro", GetProfileString("JOBST", "PathJOBST", "C:\\JOBST", "DRAFIX.INI") + "\\ARM" );
@BOD_IN.D ;

   
