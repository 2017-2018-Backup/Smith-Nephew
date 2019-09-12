// File Name:	CL_ERROR.D
// Purpose:	Save a CAD file to the directory
//		ERRORSAV
//		
//		
// Version:	1.01
// Date:		
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
//
STRING 	
	sDocFile, sFile, sDrive, sPath, sExt, sPathJOBST,
	sTmp ;
	
// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

// Use standard DRAFIX menu save (will use SaveAs if not previously saved)
//
   Execute("menu", "FileSave");
   
// Save To ERRORSAV
//
   GetData ("DocFile", &sDocFile);
   SplitPath (sDocFile, &sDrive, &sPath, &sFile, &sExt);
   Export ("drawing",  sPathJOBST + "\\ERRORSAV\\"  + sFile + ".CAD");

// Export again to ensure that path is retained
//
   Export ("drawing", sDrive + sPath + sFile + ".CAD");
   

