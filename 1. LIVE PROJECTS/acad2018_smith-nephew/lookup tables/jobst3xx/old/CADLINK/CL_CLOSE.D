// File Name:	CL_CLOSE.D
// Purpose:	Creates a .WMF file on close
//
//		
//		
// Version:	1.01
// Date:		26.Jun.95
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
//
STRING 	
	sDocFile, sFile, sDrive, sPath, sExt,
	sTmp ;
BOOL	
	FileChanged;
	
NUMBER	
	nResponse;

// Get drawing variables
//
   GetData("FileChanged", &FileChanged);
   GetData ("DocFile", &sDocFile);
   if (FileChanged){
   	nResponse = Display ("message",
   			     "YNCquestion", 
   			     sDocFile +
   			     " has been changed. Save file before closing?");
	if (nResponse == %cancel) Exit(%ok,"") ;
	if (nResponse == %yes) {   
		// Save to file using standard DRAFIX menu command	
		// defaults to "FileSaveAs" if a new drawing		     
		// We use file change to capture the odd occasion when
		// the user cancels the "FileSaveAs" dialogue box
		Execute("menu", "FileSave");
		GetData("FileChanged", &FileChanged);;
		if (FileChanged) Exit(%ok,"") ;
		
		// Get the document file name
   		GetData ("DocFile", &sDocFile);
   		SplitPath (sDocFile, &sDrive, &sPath, &sFile, &sExt);
   		
   		// Save WMF (Windows Meta File) version
   		Export ("drawing", sDrive + sPath + sFile + ".WMF", "wmf");
                Execute("menu", "FileSaveAs");
		SetData("FileChanged", 0);	
    		}	
	else
		// Don't allow DRAFIX to take over by setting changed to 0 
		 SetData("FileChanged", 0);	
	}
   
   Execute("menu", "FileClose");

