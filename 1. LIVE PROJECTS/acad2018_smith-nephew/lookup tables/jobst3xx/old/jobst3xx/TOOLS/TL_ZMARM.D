// File:		TL_ZMARM.D
// Purpose:	Insert a Arm Medial Zipper
// Version:	1.1
// Date:		9.Feb.95
// Author:		Gary George/Gerra Dunne
//		C-Gem Software/Jobst Ireland ltd
//
XY
	xyZipperEnd,
	xyPrev,
	xyNext,
	xyProfileEnd,
	xyMarkerAtStart,
	xyMarkerAtEnd,
	xyTextPos;

NUMBER
	nn,
	nL, nA,	
	nDiff,
	nLength,
	nActualLen,
	nFiguredLen;

STRING
	sFiguredLen;

HANDLE
	hEnt;

ANGLE
	aAngle;
// Get the Arm Profile Curve

   hEnt=GetEntity("Select Arm Profile","curve");

// Get End Point of Zipper
// 
   Execute("menu","SnapOff");
   GetUser("xy","Select End Point of Zipper",&xyZipperEnd);

// Calculate Zipper Length
//
   nActualLen=0;
   GetVertex(hEnt,1,&xyPrev);

// Set first marker position
//
   xyMarkerAtStart = xyPrev;
// Error Checking
//
   GetVertex(hEnt,GetVertexCount(hEnt),&xyProfileEnd);

   nDiff=xyZipperEnd.x-xyMarkerAtStart.x;
   if(nDiff<1)
	Exit(%ok,"Too Small To Draw!");
   nDiff=xyZipperEnd.x-xyProfileEnd.x;
   if(nDiff>0)
	Exit(%ok,"Zipper can not extend past end of profile!");

// Loop through curve calculating length
// 

   nn=2;
   while(nn<GetVertexCount(hEnt)){
	GetVertex(hEnt,nn,&xyNext);
   	if(xyNext.x>xyZipperEnd.x)
		break;
   	nLength=Calc("Length",xyPrev,xyNext);
	nActualLen=nActualLen+nLength;
	xyPrev=xyNext;
	nn=nn+1;
	}

//   Calculate remaining zipper length
//   Using 
//			* 
//         	                           *	*	COS (aAngle) = nL / nA
//		       * 	*
//  	     nL        *	*   
//  	             *		*
//        	        *		*
//    	    *aAngle	*
// 	  *  *  *  *  *  *  *  * *	*
// 	  |          nA	|
// 	  |		|
//xyPrev.x>|  xyZipperEnd.x>	|
//    

// Get Adjacent side length
//
   nA=xyZipperEnd.x-xyPrev.x;

// Get angle
//
   aAngle=Calc("angle",xyPrev,xyNext);

// Calculate Hyp.
//
   nL=nA/cos(aAngle);

// Length Calculation Complete
//
   nActualLen=nActualLen+nL;
   nFiguredLen=nActualLen/0.92;
   sFiguredLen=Format("length",nFiguredLen)+" Medial Zipper";
   
// Set last marker position
//
   xyMarkerAtEnd = CalcXY("relpolar", xyPrev, nL, aAngle) ;
   
// Set to Layer Notes
//
   Execute("menu","SetLayer",Table("find","Layer","Notes"));

// Insert Marker
// Markers
//	Type="closed arrow"
//	Length=  0.5inch
//	Width=    0/125inch
//

// Start of Zipper, Marker
//
   AddEntity("marker","closed arrow",xyMarkerAtStart,0.5,0.125,315);
   
// End of Zipper, Marker
//
   AddEntity("marker","closed arrow",xyMarkerAtEnd,0.5,0.125,240);

// Calculate Text Insertion Position
//	
   xyTextPos.x=xyMarkerAtStart.x+(nActualLen/2);
   xyTextPos.y=xyMarkerAtStart.y-0.25;

// Set Text Justification and Height
//
   SetData("TextVertJust",32);
   SetData("TextHorzJust",2);
   SetData("TextHeight",0.125);
   
// Insert Text
//
   AddEntity("text",sFiguredLen,xyTextPos);

// Reset to Layer 1
//
   Execute("menu","SetLayer",Table("find","Layer","1"));

	
	



