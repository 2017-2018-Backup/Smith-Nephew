// File Name:	TITLEOFF.D
// Purpose:	To Switch Title layers off only  for plotting
//		
// Version:	1.01
// Date:		20.Jan.94
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
//---------------------------------------------------------------------------------------------
//

// Local Variables 
HANDLE
	hMasked,
	hLayer;

// NB.       Explicit conditional statement  "hLayer > %zero && hLayer != 32768"
//	Because it don't work otherwise.

   UserSelection("clear");

// Make 1 the current layer.
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;

// Switch off the following layers
//	1. titlebox

   hMasked = Table("find","state", "masked") ;

   hLayer = Table("find", "layer", "titlebox");
   if (hLayer > %zero && hLayer != 32768)
	SetTableData( "layer", hLayer, "state", hMasked) ;
   hLayer = Table("find", "layer", "Data");
   if (hLayer > %zero && hLayer != 32768)
	SetTableData( "layer", hLayer, "state", hMasked) ;
// Refresh screen
   Execute("menu", "ViewRedraw") ;