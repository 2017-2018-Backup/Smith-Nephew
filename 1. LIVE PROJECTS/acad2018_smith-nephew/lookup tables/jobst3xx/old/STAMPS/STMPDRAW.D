// File Name:	STMPDRAW.D
// Purpose:	Called by VB created Draw.d macro.
//		Insert into drawing, prompting user for the position.
//		Inserts either Double or Single lines.
//
// Version:	1.2
// Date:		04.Aug.94
// Author:		Gary George
//
// Note:-
//	This macro is called from a macro created by Visual Basic
//	VB provides data in the STRING variables sStamp1, sStamp2.
//
//
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
// 2.Jun.95	gg	User Selection Clear added.
//			Exit on use of right button.

// Clear user selection
//
   UserSelection ("clear");
   UserSelection ("update");
   Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));
   Execute ( "menu", "SetColor", Table( "find",  "color", "bylayer"));

HANDLE
	hDDE,
	hLayerNotes,
	hLayerOriginal;
	
// Declared and Set in the VB created calling Macro
// STRING	 
//	sStamp1,
//	sStamp2;

XY
	xyStamp;

NUMBER
	nTextHt,
	nTextVertJustOriginal,
	nTextHorizJustOriginal,
	nTextHtOriginal;

// Procedure to set Horizontal and Vertical justification
// reduces code.
//
   procedure PRJustification () {
   	   SetData("TextHorzJust", %1);	
	   SetData("TextVertJust", %2);		
   }
   
// Get position
//
   if(!GetUser ("xy", "Stamp Position", &xyStamp)) Exit(%ok,"No Stamp Position Given!");

// Initialise, Set :-
//		Layer
//		Text Justifications 
//		Text Ht 
//
// Store the values for restoration on completion

   GetData( "LayerNumber", &hLayerOriginal) ;
   
   hLayerNotes = Table("find", "layer", "Notes") ;
   if ( hLayerNotes  != %badtable)
 	Execute ("menu", "SetLayer",  hLayerNotes) ;

   GetData("TextVertJust", &nTextVertJustOriginal);	
   GetData("TextHorzJust", &nTextHorizJustOriginal);		
   GetData("TextHeight", &nTextHtOriginal);

   PRJustification ( 2, 32) ;	// Centre & Bottom
   nTextHt = 0.1 ;
   SetData("TextHeight", nTextHt );

// Put stamp
//
   AddEntity ("text", sStamp1, xyStamp);
   if (StringLength (sStamp2) > 0) {
	xyStamp.y = xyStamp.y - nTextHt * 1.5 ;
	AddEntity ("text", sStamp2, xyStamp);
	}
	
// Reset to original values
//
   Execute ("menu", "SetLayer",  hLayerOriginal) ;
   PRJustification (nTextHorizJustOriginal, nTextVertJustOriginal) ;
   SetData("TextHeight", nTextHtOriginal);
