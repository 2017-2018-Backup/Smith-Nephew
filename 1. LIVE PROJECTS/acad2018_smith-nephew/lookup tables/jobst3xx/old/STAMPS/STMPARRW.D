// File Name:	STMPARRW.D
// Purpose:	Called by VB created Draw.d macro
//		Insert into drawing, prompting user for the position
//		Inserts either Double or Single lines
//		Inserts an Arrow and leader line
//
// Version:	2.1
// Date:		17.Aug.94
// Author:		Gary George
//
// Note:-
//	This macro is called from a macro created by Visual Basic
//	VB provides data in the STRING variables sStamp1, sStamp2.
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
	xyArrow,
	xyStamp;
NUMBER
	nTextHt,
	nCase,
	nTextVertJustOriginal,
	nTextHorizJustOriginal,
	nTextHtOriginal;
ANGLE
	aAngle ;

// Procedure to set Horizontal and Vertical justification
// reduces code.
//
   procedure PRJustification () {
   	   SetData("TextHorzJust", %1);	
	   SetData("TextVertJust", %2);		
   }
	
// Get Arrow start and text position
//
   if(!GetUser ("xy", "Arrow Position", &xyArrow))   Exit(%ok,"No Arrow Position Given");
   if(!GetUser ("xy", "Stamp Position", &xyStamp)) Exit(%ok,"No Stamp Position Given");

// Initialise, Set :-
//		Layer
//		Text Justifications 
//		Text Ht 
//
// Store the values for restoration on completion

   GetData( "LayerNumber", &hLayerOriginal) ;
   
   hLayerNotes = Table("find", "layer", "Notes") ;
   if ( hLayerNotes > %zero && hLayerNotes != 32768)
 	Execute ("menu", "SetLayer",  hLayerNotes) ;

   GetData("TextVertJust", &nTextVertJustOriginal);	
   GetData("TextHorzJust", &nTextHorizJustOriginal);		
   GetData("TextHeight", &nTextHtOriginal);

   nTextHt = 0.1 ;
   SetData("TextHeight", nTextHt );


// Calculate arrow and text justification
//
   aAngle = Calc ("angle", xyStamp, xyArrow);
   if ( (aAngle >= 0 && aAngle <= 45) || aAngle > 315) {
	PRJustification( 4, 16);
	nCase = 1;
	}
   if (aAngle > 45  && aAngle <= 135)  {
	PRJustification( 2, 8);
	nCase = 2;
	}

   if (aAngle > 135 && aAngle <= 225){
	PRJustification( 1, 16);
	nCase = 3;
	}
   if (aAngle > 225 && aAngle <= 315){
	PRJustification( 2, 32);
	nCase = 4;
	}

// Put Arrow and line
//
   AddEntity ("marker", "closed arrow", xyArrow, 0.5 , 0.125, aAngle+180 ) ;
   AddEntity ("line", xyStamp, xyArrow) ;

// Put stamp text, nb Special case for double lines
//
   if (StringLength (sStamp2) == 0) {
  	AddEntity ("text", sStamp1, xyStamp);
	}
   else {
	if (nCase == 1 || nCase == 3 ){
		xyStamp.y = xyStamp.y + nTextHt * 0.75 ;
	  	AddEntity ("text", sStamp1, xyStamp);
		xyStamp.y = xyStamp.y - nTextHt * 1.5 ;
		AddEntity ("text", sStamp2, xyStamp);
		}
	if (nCase == 2 ){
	  	AddEntity ("text", sStamp1, xyStamp);
		xyStamp.y = xyStamp.y - nTextHt * 1.5 ;
		AddEntity ("text", sStamp2, xyStamp);
		}
	if (nCase == 4 ){
		AddEntity ("text", sStamp2, xyStamp);
		xyStamp.y = xyStamp.y + nTextHt * 1.5 ;
	  	AddEntity ("text", sStamp1, xyStamp);
		}
	}

// Reset to original values
//
   Execute ("menu", "SetLayer",  hLayerOriginal) ;
   PRJustification (nTextHorizJustOriginal, nTextVertJustOriginal) ;
   SetData("TextHeight", nTextHtOriginal);
