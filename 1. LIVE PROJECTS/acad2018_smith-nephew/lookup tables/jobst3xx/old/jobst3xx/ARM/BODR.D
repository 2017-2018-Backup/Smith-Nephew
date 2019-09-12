// File Name:	BODR.D
// Purpose: 	Shell for RIGHT Sleeve of a body suit
//		Calls BOD_IN.D 
// Version:	1.1
// Date:		10.Sept.97
// By:		Gary George
STRING	sSleeve;
sSleeve = "Right";
SetData("PathMacro", GetProfileString("JOBST", "PathJOBST", "C:\\JOBST", "DRAFIX.INI") + "\\ARM" );
@BOD_IN.D ;

   
