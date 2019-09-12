// Create Curve datafile
// For Sleeveless body suit


HANDLE
	hEnt, hFile;
XY	
	xyPt1, xyPt2;
NUMBER		
	aAngle, nLength, ii, nVertexCount,nPrevAngle  ;
STRING 
	sLine;

GetUser ("xy", "point", &xyPt1) ;
nPrevAngle = 0 ;
hFile = Open ("file", "E:\\JOBST\\JOBST2xx\\TEMPLTS\\BODYCURV.DAT", "readonly") ;
if (hFile) {
StartPoly("polyline") ;

   	while ( GetLine(hFile, &sLine)) {
   		ScanLine (sLine, "blank", &nLength, &aAngle) ;
	     	nPrevAngle=nPrevAngle + aAngle;
   		
   		xyPt2 = CalcXY("relpolar",  xyPt1 , nLength,  nPrevAngle ) ;
   		
   		AddVertex (xyPt1) ;
	     	xyPt1 = xyPt2 ;
	
		}
	Close("file", hFile) ;
EndPoly();
}