// File Name:	BD_BOTH.D
// Purpose:		Draws a template for the leg portion of the
//			bodysuit 
//	
// Version:		1.01
// Date:		9.Oct.97
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
// set template defaults
// Declare variables
XY		xyPt1, xyTmp;
NUMBER	ii, jj, nMinHt, nMaxHt, nReduction, nX, nSeam, n20Len, nTickHt, nTickStep, nMaxLength, nMinLength;
STRING	sSymbol, sMinHtColour, sMaxHtColour ;
		

// Set Symbol Library
   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB");

// Set layer   
   hLayer = Table("find", "layer", "Construct") ;
   if ( hLayer <> %badtable ) Execute ("menu", "SetLayer",  hLayer) ;

// Seam size
   nSeam = .25;

procedure PR_DrawTapeScale() {
   	xyPt1 = xyStart ;
	xyPt1.x = xyPt1.x + nX ;
	xyTmp.x = xyPt1.x ;
	
	nMaxLength = n20Len / 20 * nMaxHt ; 
	xyTmp.y = xyStart.y + nSeam + nMaxLength ;
	hEnt = AddEntity ("marker", "xmarker", xyTmp, 0.2, 0.2) ;
	SetEntityData (hEnt, "color",  Table("find", "color", sMaxHtColour ));

	nMinLength = n20Len / 20 * nMinHt ; 
	if (nMinHt != 0 ) {
		xyTmp.y = xyStart.y + nSeam + nMinLength ;
		hEnt = AddEntity ("marker", "xmarker", xyTmp, 0.2, 0.2) ;
		SetEntityData (hEnt, "color",  Table("find", "color", sMinHtColour ));
		}	
	else {
		nMinLength = nMaxLength ; 
		}
	
	// Add Ticks and Text at each scale
	// Start ticks 2 below and 2 above
	ii = int ( nMinLength / (n20Len / 20)) - 1 ;
          jj = (int ( nMaxLength / (n20Len / 20)) - 1) + 2 ;

	nTickStep = n20Len / (20 * 8);
	nTickHt = xyStart.y + nSeam + ii * nTickStep * 8;

	// Add Text
	SetData("TextHeight", 0.125);
	SetData("TextHorzJust", 1);
	
	// Reduction text	
	AddEntity ("text", MakeString("long", nReduction), xyPt1.x + 0.2, nTickHt - nSeam*3) ;
	
	// Maximum text	
	hEnt = AddEntity ("text", Format("length", nMaxHt), xyPt1.x + 0.2, nTickHt - nSeam) ;
	SetEntityData (hEnt, "color",  Table("find", "color", sMaxHtColour ));
	
	// Minimum text	
	if ( nMinHt != 0 && nMinHt != nMaxHt) {
		hEnt = AddEntity ("text", Format("length", nMinHt), xyPt1.x + 0.2, nTickHt - nSeam*2) ;
		SetEntityData (hEnt, "color",  Table("find", "color", sMinHtColour ));
		}

	// Add tape ID
   	if ( !Symbol("find", sSymbol)) Exit(%cancel, "Can't find a symbol to insert\nCheck your installation, that JOBST.SLB exists");
	AddEntity("symbol", sSymbol,  xyPt1.x, nTickHt ) ;

	// Add graticules
	while (ii <= jj ) {
		nTickLength = 0.22;
		AddEntity( "line", xyPt1.x, nTickHt, xyPt1.x + nTickLength, nTickHt) ;
		// Add Tick Text
		if ( ii < 10)  SetData("TextHorzJust", 4);
			else SetData("TextHorzJust", 2);
		AddEntity("text", MakeString("long",ii), xyPt1.x + nSeam, nTickHt - 0.01, 0.06, 0.1, 90 ) ; 
		nTickLength = 0.05;
		AddEntity( "line", xyPt1.x, nTickHt + nTickStep, xyPt1.x + nTickLength, nTickHt + nTickStep) ;
		AddEntity( "line", xyPt1.x, nTickHt + nTickStep*3, xyPt1.x + nTickLength, nTickHt + nTickStep*3) ;
		AddEntity( "line", xyPt1.x, nTickHt + nTickStep*5, xyPt1.x + nTickLength, nTickHt + nTickStep*5) ;
		AddEntity( "line", xyPt1.x, nTickHt + nTickStep*7, xyPt1.x + nTickLength, nTickHt + nTickStep*7) ;
		nTickLength = 0.08;
		AddEntity( "line", xyPt1.x, nTickHt + nTickStep*2, xyPt1.x + nTickLength, nTickHt + nTickStep*2) ;
		AddEntity( "line", xyPt1.x, nTickHt + nTickStep*6, xyPt1.x + nTickLength, nTickHt + nTickStep*6) ;
		nTickLength = 0.12;
		AddEntity( "line", xyPt1.x, nTickHt + nTickStep*4, xyPt1.x + nTickLength, nTickHt + nTickStep*4) ;
		ii = ii + 1;
		nTickHt = xyStart.y + nSeam + ii * nTickStep * 8;
		}
} 

