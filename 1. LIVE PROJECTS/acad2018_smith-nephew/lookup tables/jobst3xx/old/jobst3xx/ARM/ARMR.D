// File Name:	ArmR.D
// Purpose: 	Shell for RIGHT Sleeve
//		Calls ARM_IN.D 
// Version:	1.1
// Date:		22.Mar.95
// By:		Gary George
STRING	sSleeve;
sSleeve = "Right";
SetData("PathMacro", GetProfileString("JOBST", "PathJOBST", "C:\\JOBST", "DRAFIX.INI") + "\\ARM" );
@ARM_IN.D ;

   
