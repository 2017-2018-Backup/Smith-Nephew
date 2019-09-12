// File Name:  CL_POST.D  <Dir> <Dwg.CAD> <WO.TXF>
// Purpose:     To Link TRITON and JOBST CAD System.
//                     Drawings will have been created post-imageABLE 
//                     or they will be new drawings.
//
// Version:        1.01
//
// Date:	    08.Jun.95
//
// Author:	    Gary George
//
// Arguments:   <Dir>            Directory where final drawing is to be stored
//                      <Dwg.CAD>  Original Drawing (Including full path)
//                      		           or
//                      		     NEW to indicate a new drawing.                          
//                      <WO.TXF>    Triton eXchange File (Including full path)
//              
//
// Note
//      Because DRAFIX is so bad at handling errors.
//      No Error checking is possible on the validity of the 
//      arguments.
//      Therefor, these arguments  >>> MUST <<< be valid
//      and checked by the invoking procedures.
//
//---------------------------------------------------------------------------------------------
// REVISIONS:   
// Date         By      Action
//---------------------------------------------------------------------------------------------
//

STRING 
	sTmp,
	sOldDrawing,
	sWO_TXF,
	sWorkOrder,
	sOldWorkOrder,
	sDrawingDir;
	
  
   Display ("cursor", "wait", "TRITON to CAD Link");
	
// Get Arguments
// See NOTE above.
//
   sDrawingDir  = GetArgument (1) ;
   // Append backslash "\", But ONLY if sDrawingDir dosn't end in "\" 
   if (!StringCompare("\\", StringMiddle(sDrawingDir  ,StringLength(sDrawingDir),1))) sDrawingDir = sDrawingDir   + "\\";       

// Get the existing drawing
// Split out the old workorder.
//            
   sOldDrawing = GetArgument (2) ;
   SplitPath( sOldDrawing, &sTmp, &sTmp, &sOldWorkOrder,  &sTmp);
 
// From the work order TXF file get the implicit work order number.
// ie T:\MISC\123456.TXF => WO# == 1234546
// 
   sWO_TXF = GetArgument (3) ; 
   SplitPath( sWO_TXF, &sTmp, &sTmp, &sWorkOrder, &sTmp);

// if New   THEN Create DATA Base Fields and Layers
//      ELSE Import existing drawing.
   
   SetData("PathMacro", GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") );
   
   if (StringCompare("NEW", sOldDrawing)){
	Execute("menu", "FileNew") ;
	@JOBSTNEW.D ;
  	}
   else
	if (!Import ("drawing", sOldDrawing)) Exit (%abort, "Can't import Old DRAWING!");  


// Store TXF file and Work Order in the UserCharN for later use by 
// TRITONtoCAD VB programme 
//
// UserChar1 = <WO.TXF> Triton eXchange File (Including full path)
// UserChar2 = WorkOrder        WO stripped from <WO.TXF> above
// UserChar3 = Message to CL_TXF either:-
//		"NEW"		- empty drawing
//		"11111.CAD"    	- i.e. uses existing drawing
//		"INSPECTION" 	- set below
//

   SetData ( "UserChar1", sWO_TXF) ;
   SetData ( "UserChar2", sWorkOrder) ;
   SetData ( "UserChar3", sOldDrawing) ;

// Save Drawing (Only if sWorkOrder <> sOldWorkorder)
// IE save only if the work orders are different.
//     otherwise we assume that imageABLE is supplying for
//     the inspection stage
   if (!StringCompare(sWorkOrder, sOldWorkOrder))
	Export ("drawing", sDrawingDir + sWorkOrder + ".CAD", "cad") ;
   else	{
   	// Export .WMF file so that imageABLE can pick it up
      	// This will be identical to the one created previously
      	//
      	Export ("drawing", sDrawingDir + sWorkOrder  + ".WMF", "wmf");
      	
	// Export again to ensure that it is the CAD version that is being worked
	// on.  Otherwise we end up using the .WMF version.
	//
	Export ("drawing", sDrawingDir + sWorkOrder + ".CAD", "cad") ;
	
	// Let CL_TXF know that it is for inspection
	//
   	SetData ( "UserChar3", "INSPECTION") ;
   	
      	// Fool DRAFIX into thinking that no changes have been made
      	//
  	SetData("FileChanged", 0);	
	}
    	
// Start the VB Programme to display and manipulate the TRITON supplied
// patient and work order data
// 
  Document("maximize");
  Execute ("application", "jobstart TRITONtoCAD", "normal") ;
  
// End CL_POST.D  
  
  