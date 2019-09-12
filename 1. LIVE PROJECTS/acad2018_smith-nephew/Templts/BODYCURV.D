// Create Curve datafile
// For Sleeveless body suit
// 
//AddEntity("poly", "fitted",
//0,0,
//0.25,  	0.4,
//0.5,   	0.62,
//1.0,   	0.9,
//1.5,   	1.08,
//2.0,	1.17,
//3.0,	1.27,
//4.0,	1.26,
//5.0,	1.19,
//6.0,	1.08,
//7.0,	0.92,
//8.0,	0.74,
//9.0,	0.56,
//10.0,	0.35,
//11.0,	0.15
//) ;



HANDLE
	hEnt, hFile;
XY	
	xyPt1, xyPt2;
NUMBER		
	aAngle, nLength, ii, nVertexCount ;

hEnt = GetEntity ("select polyline", "polyline") ;

hFile = Open ("file", "E:\\JOBST\\JOBST2xx\\TEMPLTS\\BODYCURV.DAT", "write") ;
	ii = 1;
	nVertexCount = GetVertexCount(hEnt) ;
	GetVertex (hEnt, ii, &xyPt1) ;
	ii = 2;
	aPrevAngle = 0 ;
   	while ( ii <= nVertexCount) {
	     	GetVertex (hEnt, ii, &xyPt2) ;
	     	aAngle = Calc("angle", xyPt1, xyPt2) ;
	     	aDeltaAngle = aAngle - aPrevAngle;
	     	nLength = Calc("length", xyPt1, xyPt2) ;
	     	xyPt1 = xyPt2 ;
	     	aPrevAngle = aAngle;
		PrintFile(hFile,  nLength, aDeltaAngle , "\n") ;
		ii = ii + 1;
		}
	Close("file", hFile) ;
