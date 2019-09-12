// File Name:	MERGE.D <sMergeCADFile>
//
// Purpose:		Merges the current drawing with the file given
//			in the argument 
//			
// Version:		1.01
//
// Date:		06.Feb.98
//
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
// Notes:


Display ("cursor", "wait", "Loading Data");

HANDLE	
	hChan,
	hSym,
	hPrevSym,
	hDDE,
	hEnt;
	
XY	
	xySymO,
	xyMin, xyMax,
	xyInPoly1, xyInPoly2, xyInPoly3, xyInPoly4,
	xyMerge,
	xyCurrentMPD,
	xyScale;
	
ANGLE
	aAngle;

STRING	
	sSymName, sDataSymbol ,
	sLeg, sSleeve,
	TAB,
	NL, sSide ,
	sMessage,
	sMessageDuplicateData ,
	sSymSelection,
	sPathJOBST,
	sCurrentCADFile, 
	sMergeCADFile ,
	sCurrentMPD,
	sTest, sMessge,
	// Strings are initialized to empty "", this ensures empty strings
	// are poked if there is no DB field.
	s1, s2, s3, s4, s5, s6, s7, s8, s9, 
	s11, s12, s13, s14, s15, s16, s17, s18, s19;
	
NUMBER
	nDeltaX ,
	nDeltaY,
	nMPD,
   	nVestLeft,
   	nVestRight,
   	nVest ,
   	nArmLeft,
   	nArmRight,
   	nArm,
   	nGloveLeft,
   	nGloveRight,
   	nGlove,
   	nHeadNeck,
   	nLegLeft,
   	nLegRight,
   	nLeg,
   	nWaistLegLeft,
   	nWaistLegRight,
   	nWaist,
   	nSuitBody ;
   
// Functions
//
NUMBER
	nCurrentUID;
BOOL
	bInsideCurrentMPD ;
// %1 = UID of previous sysmbol
 
function FN_CheckForDuplicate () {
	if (%1 > 0) sMessageDuplicateData = sMessageDuplicateData + sSymName + TAB + sSide + NL ;

	bInsideCurrentMPD = Calc ( "inpoly", xySymO, xyInPoly1,  xyInPoly2,  xyInPoly3,  xyInPoly4) ;
	nCurrentUID = UID ("get", hSym) ;
	
	if ( !bInsideCurrentMPD) {
		if ( %1 > 0){
			SetGeometry(hSym, sSymName, xySymO.x  + nDeltaX, xySymO.y + nDeltaY - 2.375, xyScale, aAngle);
			}
		else{
			SetGeometry(hSym, sSymName, xySymO.x  + nDeltaX, xySymO.y + nDeltaY, xyScale, aAngle);
			}
		}
	else {
		if ( %1 > 0) {
			if ( hPrevSym = UID("find", %1)) {
				GetGeometry(hPrevSym, &sSymName, &xySymO, &xyScale, &aAngle);
				SetGeometry(hPrevSym, sSymName, xySymO.x, xySymO.y - 2.375, xyScale, aAngle);
				}			 		
			}
		}
	
	return (nCurrentUID) ;	// Return the UID of the current symbol found
	
} // End function

// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   TAB = "\t";
   NL = "\n" ;
   
// Get the  "mainpatientdetails" of current file
//

   hChan = Open ("selection",	"DB SymbolName = 'mainpatientdetails'"  );  
   if (hChan) {
      	ResetSelection(hChan);      	
 	// Get mainpatientdetails data base fields
	//
          while (hSym = GetNextSelection(hChan)) {
   		GetGeometry(hSym, &sSymName, &xySymO, &xyScale, &aAngle);
      		if(StringCompare(sSymName, "mainpatientdetails" )){ 
			GetDBData( hSym, "fileno", &s1);
			GetDBData( hSym, "patient", &s2); 
 			GetDBData( hSym, "age",  &s3);
			GetDBData( hSym, "units", &s4); 
			GetDBData( hSym, "sex", &s5);
			GetDBData( hSym, "Diagnosis", &s6); 
			GetDBData( hSym, "WorkOrder", &s7); 	
	   		GetDBData( hSym, "orderdate", &s8);
   			GetDBData( hSym, "TemplateEngineer", &s9);
   			sCurrentMPD = s1 + s2 + s3 + s4 + s5 + s6 + s7 + s8 + s9;
   			xyCurrentMPD = xySymO;
   			GetExtent (hSym, xyInPoly1 , xyInPoly3) ;
   			xyInPoly1.y = xyInPoly1.y - 0.25 ;
   			xyInPoly1.x = xyInPoly1.x - 0.25 ;
   			xyInPoly2.y = xyInPoly1.y  ;
   			xyInPoly2.x = xyInPoly3.x  ;
   			xyInPoly4.y = xyInPoly3.y  ;
   			xyInPoly4.x = xyInPoly1.x  ;   			
   			}
   		}
	}
   
   Close ( "selection", hChan); 

// Get sheet size so that we can establish an insertion point for the 
// merge drawing   
   
   GetData("DrawingSheet", &xyMin, &xyMax) ;
   xyMerge.x = xyMax.x ;
   xyMerge.y = xyMin.y ;
   
// Merge the drawing
//
   sMergeCADFile = GetArgument (1) ;
   
//   Print (sMergeCADFile) ;

   if ( ! StringLength (sMergeCADFile) > 0 ) Exit(%cancel, "No CAD file given to merge!");
   if ( ! Merge("drawing", sMergeCADFile , "cad", xyMerge) )Exit(%cancel, "Unable to complete merge!");
   UserSelection ("clear" );
   UserSelection ("update") ;
   
// Sort out the fact that we will have more than one mainpatientdetails 
// Delete the one we don't want
//
   hChan = Open ("selection", "DB SymbolName = 'mainpatientdetails'"  );  
   if (hChan) {
      	ResetSelection(hChan);      	
 	// Get mainpatientdetails data base fields
	//
          while (hSym = GetNextSelection(hChan)) {
         		GetGeometry(hSym, &sSymName, &xySymO, &xyScale, &aAngle);
           		         		
   		if(StringCompare(sSymName, "mainpatientdetails" )) {
   			nMPD = nMPD + 1 ;
			GetDBData( hSym, "fileno", &s11);
			GetDBData( hSym, "patient", &s12); 
 			GetDBData( hSym, "age",  &s13);
			GetDBData( hSym, "units", &s14); 
			GetDBData( hSym, "sex", &s15);
			GetDBData( hSym, "Diagnosis", &s16); 
			GetDBData( hSym, "WorkOrder", &s17); 	
	   		GetDBData( hSym, "orderdate", &s18);
   			GetDBData( hSym, "TemplateEngineer", &s19);
   			sTest = s11 + s12 + s13 + s14 + s15 + s16 + s17 + s18 + s19;
   			if (!StringCompare( sTest, sCurrentMPD)) {
   				sMessage = "" ;
   				DeleteEntity (hSym) ;
   				if (!StringCompare(s1,s11)) sMessage = sMessage + "File" + TAB + s1 +TAB + s11 + NL ;
   				if (!StringCompare(s2,s12)) sMessage = sMessage + "Name"  + TAB + "'" + s2 + "'" + TAB + "'" + s12 + "'" + NL ;
   				if (!StringCompare(s3,s13)) sMessage = sMessage + "Age" + TAB + s3 +TAB + s13 + NL ;
   				if (!StringCompare(s4,s14)) {
   								     sMessage = sMessage + "Units" + TAB + s4 +TAB + s14 + NL ;
   								     sMessage = sMessage +  TAB + " WARNING! >>>> Units differ <<<< WARNING!" + NL ;
   								     }
   				if (!StringCompare(s5,s15)) sMessage = sMessage + "Sex" + TAB + s5 +TAB + s15 + NL ;
   				if (!StringCompare(s6,s16)) sMessage = sMessage + "Diag." + TAB + s6 +TAB + s16 + NL ;
   				if (StringLength (sMessage) >0) sMessage = "The following differences were found between the current patient details\nand the patient details from the merged drawing file." + NL +  NL + sMessage ;
   				nDeltaX = xyCurrentMPD.x - xySymO.x ;
   				nDeltaY = xyCurrentMPD.y - xySymO.y ;
   				}
   			}
   			
   		}
	}
	
   Close ("selection", hChan) ;
  	
// Sort out the fact that we will have more than one mainpatientdetails and possibly duplicate symbols 
//
   sMessageDuplicateData = "" ;
   
   hChan = Open ("selection",	"type = 'symbol'" );  
   if (hChan) {
      	ResetSelection(hChan);      	
 	// Get mainpatientdetails data base fields
	//
          while (hSym = GetNextSelection(hChan)) {
         		GetGeometry(hSym, &sSymName, &xySymO, &xyScale, &aAngle);
         		
         		sLeg = "" ; sSleeve = "" ;
		GetDBData( hSym, "sleeve", &sSleeve);
		GetDBData( hSym, "leg", &sLeg); 		
		sSide = sLeg + sSleeve ;		  	// We can do this as 'leg' and 'sleeve' are
         		sDataSymbol = sSymName + sSide  ;	// never used together.
         		         		   			
   		if(StringCompare(sDataSymbol , "vestarmLeft" ))  nVestLeft = FN_CheckForDuplicate( nVestLeft  );
   		if(StringCompare(sDataSymbol , "vestarmRight" ))  nVestRight = FN_CheckForDuplicate( nVestRight  );
   		if(StringCompare(sDataSymbol , "vestbody" )) nVest = FN_CheckForDuplicate( nVest  );
   		if(StringCompare(sDataSymbol , "armarmLeft" ))  nArmLeft = FN_CheckForDuplicate( nArmLeft );
   		if(StringCompare(sDataSymbol , "armarmRight" ))  nArmRight = FN_CheckForDuplicate( nArmRight  );
   		if(StringCompare(sDataSymbol , "armcommon" ))  nArm = FN_CheckForDuplicate( nArm  );
   		if(StringCompare(sDataSymbol , "glovegloveLeft" )) nGloveLeft = FN_CheckForDuplicate( nGloveLeft  );
   		if(StringCompare(sDataSymbol , "glovegloveRight" )) nGloveRight = FN_CheckForDuplicate( nGloveRight  );
   		if(StringCompare(sDataSymbol , "glovecommon" )) nGlove = FN_CheckForDuplicate( nGlove );
   		if(StringCompare(sDataSymbol , "headneck" )) nHeadNeck = FN_CheckForDuplicate( nHeadNeck );
   		if(StringCompare(sDataSymbol , "leglegLeft" )) nLegLeft = FN_CheckForDuplicate( nLegLeft );
   		if(StringCompare(sDataSymbol , "leglegRight" )) nLegRight = FN_CheckForDuplicate( nLegRight  );
   		if(StringCompare(sDataSymbol , "legcommon" )) nLeg = FN_CheckForDuplicate( nLeg  );
   		if(StringCompare(sDataSymbol , "waistlegLeft" )) nWaistLegLeft = FN_CheckForDuplicate( nWaistLegLeft );
   		if(StringCompare(sDataSymbol , "waistlegRight" )) nWaistLegRight = FN_CheckForDuplicate( nWaistLegRight );
   		if(StringCompare(sDataSymbol , "waistbody" )) nWaist = FN_CheckForDuplicate( nWaist );
   		if(StringCompare(sDataSymbol , "suitbody" )) nSuitBody = FN_CheckForDuplicate( nSuitBody );
   		}
	}
	
   Close ("selection", hChan) ;

// Display messages
// 
   if (StringLength (sMessageDuplicateData ) >0 ) {
   	sMessageDuplicateData = "The following duplicate Data definitions were found.\n" + NL  + sMessageDuplicateData ;
	// Set Symbol Library
   	SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB");
   	Execute ("menu", "SetLayer",  Table("find", "layer", "titlebox")) ;
   	if ( Symbol("find", "duplicatedatasymbols")) AddEntity("symbol",  "duplicatedatasymbols",  xyCurrentMPD) ;
  	}
  	
   if ( nMPD > 2 ) sMessageDuplicateData = sMessageDuplicateData  + "Warning! More than 2 mainpatientdetails found.\n"  + NL ;

   if (StringLength (sMessage) >0 || StringLength (sMessageDuplicateData )> 0 ) {
   	hChan = Open ("file", "C:\\JOBST\\ERRORS.DAT", "write") ;
   	PrintFile(hChan, sMessage, "\n");
   	PrintFile(hChan, sMessageDuplicateData, "\n");
   	Close ("file", hChan) ;
	Display("window",  "C:\\JOBST\\ERRORS.DAT", "text", "Merging error messages", "");
   	}
   
// Reset to layer 1
//   	
   Execute ("menu", "SetLayer",  Table("find", "layer", "titlebox")) ;
