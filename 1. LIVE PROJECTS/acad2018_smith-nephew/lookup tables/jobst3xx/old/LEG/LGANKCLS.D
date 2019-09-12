// File Name:	LGANKCLS.D
// Purpose:	Draw closing lines for Ankle
//
// Version:	1.01
// Date:		25.Jan.94
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

   Execute ("menu", "SetLayer", hTemplateLayer);
   if (SmallHeel) 
	xyO.x = xyO.x - 0.375;
   else
	xyO.x = xyO.x - 0.75;
   
   xyPt1 = CalcXY("relpolar", xyProfileLast, 0.25, 270) ;
   AddEntity ("line", xyPt1, xyProfileLast) ;
   hEnt = AddEntity ("line", xyPt1, xyO) ;	
   SetDBData(hEnt, "ID", sLegStyle+sFileNo+sLeg+"ClosingLine") ;

   if (xyOtemplate.x < xyToeSeam.x)
  	 AddEntity ("line" , xyO, xyToeSeam.x, xyOtemplate.y);
   else
  	 AddEntity ("line" , xyO, xyOtemplate);
   
// Seam TRAM Lines
   PRSetLayerNotes() ;
   AddEntity("line",xyToeSeam.x, xyO.y + nSeam + 0.5, xyO.x, xyO.y + nSeam + 0.5);
   AddEntity("line", xyToeSeam.x, xyO.y + nSeam, xyO.x, xyO.y + nSeam );