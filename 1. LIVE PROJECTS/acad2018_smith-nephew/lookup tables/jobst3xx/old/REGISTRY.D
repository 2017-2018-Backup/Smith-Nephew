// FileName:	REGISTRY.D
// Purpose:	Updates the Registry with values that will allow the
//		running of large Macros
//		
// Version:	1.01
// Author:	Gary George
// Last Date:	30 Sept 1999
//

// General:
//
// HKEY_CURRENT_USER\Software\Softdesk\Drafix\4.0

	SetRegistryString( "Macro", "SizeProgram", "12000") ;
	
	SetRegistryString( "Macro", "SizeOutput", "32000") ;

	SetRegistryString( "Macro", "InitDoc", "1") ;
	
// End
	 