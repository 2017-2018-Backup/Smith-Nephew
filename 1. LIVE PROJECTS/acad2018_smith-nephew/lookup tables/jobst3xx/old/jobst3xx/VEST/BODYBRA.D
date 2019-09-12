// File Name:	BODYBRA.D
// Purpose:	Draw BRA cup disks on body
// Version:	2.01
// Date:		02.Nov.93
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
// 10.Nov.93	gg	Fixed Bra cup = "None" bug.
//			Allow user selected disks.
// 20.Nov.93	gg	Fixed Bra Cup selection bug
//			Ref JOBST letter 10.Nov.93
// 20.Jan.93	gg	Seperate BRA disk symbols 
//			now used 
// 28.Jan.94	gg	For bra disks over 5 cut back by 1
//
// 18.Oct.95	gg	Sizing performed by Visual Basic
//			BODYBRA.D now only inserts the 
//			Bra disks
//
// 4.Jan.95	gg	Use diff. nBraFac_1 for sleeveless
//
// note:-	nBraAxillaHt has aleady been compensated for xyO.
//	It is only at this stage do we worry about the fact that the
//	cups are the same.
 
   if (StringCompare(sLtAxillaType, sRtAxillaType) && StringCompare(sLtAxillaType,"Sleeveless"))
   	nBraFac_1 = nBraFac_SleeveLess ;

   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB") ;   	     
   if ( Symbol("find",  "bradisk1")){
	if ( nDiskLt != nDiskRt) jj = 2 ;	
		else jj = 1;
	SetData("TextFont", 0);		//CAD Block
	SetData("TextAspect" , 0.6) ;
	SetData("TextVertJust", 16);		// Center
	SetData("TextHorzJust", 2);		// Center
	SetData("TextHeight", 0.125);
	while ( jj > 0 ) {
		if ( jj==2 ) {
			nDisk = nDiskLt;
 			nBraCLOffset = nBraCLOffsetLt ;
			nDiskXoff = nDiskXoffLt;
			nDiskYoff = nDiskYoffLt;
			sTmp = " (Left)";
			}
		if ( jj == 1 ) {
			nDisk = nDiskRt;
			sTmp = " ";
			nBraCLOffset = nBraCLOffsetRt ;
			nDiskXoff = nDiskXoffRt;
			nDiskYoff = nDiskYoffRt;
			}
		if ( nDisk > 0 ) {
			sDisk = "bradisk" + MakeString("long",nDisk)  ;
			if (Symbol("find",  sDisk)) {
				xyDisk.x = nBraAxillaHt - nDiskXoff - nBraFac_1 ;
				xyDisk.y = xyO.y + nBraCLOffset;
				Import("symbol", sDisk);
 				AddEntity("symbol", sDisk , xyDisk );
				}
			AddEntity("text",  "No. " + MakeString ("long",nDisk) + sTmp
	 	         	         ,nBraAxillaHt - nDiskXoff  - nBraFac_1
		        	         ,xyO.y + nBraCLOffset + nDiskYoff ) ;	
			PRSetLayerConstruct () ;
			AddEntity("line", nBraAxillaHt
				,xyO.y + nBraCLOffset
				,nBraAxillaHt - (nDiskXoff  - nBraFac_1)*2
				,xyO.y + nBraCLOffset) ;
			AddEntity("line", nBraAxillaHt - nBraFac_1
				,xyO.y + nBraCLOffset
				,nBraAxillaHt  - nBraFac_1
				,xyO.y + nBraCLOffset + (nDiskYoff) * 2 ) ;
			PRSetLayerNotes();
			}
		if ( nDisk < 0 ) {
			if ( jj == 2) 
			      AddEntity("text", "No Cup - Left"
	 	         	         ,nBraAxillaHt - nBraPosFacX   - nBraFac_1
		        	         ,xyO.y + nBraPosFacY  ) ;	
			else
			      AddEntity("text", "No Cup - Right"
	 	         	         ,nBraAxillaHt - nBraPosFacX  - nBraFac_1
		        	         ,xyO.y + nBraPosFacY + 0.5 ) ;	
			}
		jj = jj - 1 ;
		} // Endwhile
	}
   else 
	 Exit(%cancel, "Can't find BRADISK symbols to insert\nCheck your installation, that JOBST.SLB exists");
 
