// File Name:   CL_SAVE.D
// Purpose:     Explitcit save
//              Creates a .WMF file
//
//              
//              
// Version:     1.01
// Date:                26.Jun.95
// Author:              Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:   
// Date         By      Action
//---------------------------------------------------------------------------------------------
//
STRING  
	sDocFile, sFile, sDrive, sPath, sExt,
	sTmp ;

// Use standard DRAFIX menu save (will use SaveAs if not previously saved)
//
   Execute("menu", "FileSave");
   
// Save WMF (Windows Meta File) version
//
   GetData ("DocFile", &sDocFile);
   SplitPath (sDocFile, &sDrive, &sPath, &sFile, &sExt);
//   Export ("drawing", sDrive + sPath + sFile + ".WMF", "wmf");
//   Export ("drawing", "C:" + "\\IMAGE\\" + sFile + ".WMF", "wmf");

   Execute("menu","FileSaveAs");
// Export again to ensure that it is the CAD vesrion that is being worked
// on.  Otherwise we end up using the .WMF version.
//
   Export ("drawing", "C:" + "\\IMAGE\\" + sFile + ".CAD", "cad");
   Execute("menu", "FileSave");

   Export ("drawing", sDrive + sPath + sFile + ".CAD", "cad");
   Execute("menu", "FileSave");
