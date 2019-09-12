// File Name:	SlvL.D
// Purpose: 	Shell for LEFT Sleeve
//		Calls SLV_IN.D 
// Version:	1.1
// Date:		22.Mar.95
// By:		Gary George
STRING	sSleeve;
sSleeve = "Left";
SetData("PathMacro", GetProfileString("JOBST", "PathJOBST", "C:\\JOBST", "DRAFIX.INI") + "\\ARM" );
@SLV_IN.D ;

   
