// File Name:	MIRPROF.D
// Purpose:	To mirror Bodysuit back profile after user modification.
//
// Version:	1.00
// Date:	20/7/96
// Author:	Paul O'Rawe
//		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:   
// Date By      Action
//---------------------------------------------------------------------------------------------
//

HANDLE	hProfile, hCentreLine, hSelectionCL, hSelectionCurves, hNewProfile, hLayer, hLayerNotes, hEnt;
NUMBER	ii, nVertexCount, nn;
STRING	sProfileID, sCurveType, sSide;
XY		xyStart, xyEnd, xyPt1;

hProfile = GetEntity("Select Body Profile Polyline to Mirror:", "curve");
if (!hProfile) Exit(%ok, "Profile not selected") ;

if (hProfile && GetDBData(hProfile, "ID", &sProfileID)) {
		// Get the CenterLine
		//
		UserSelection ("clear");
		UserSelection ("update");
		hSelectionCL = Open("selection", "DB ID = '" + sProfileID + "' AND  DB curvetype = 'CentreLine'");
		if (hSelectionCL) {
				ResetSelection(hSelectionCL);
				hCentreLine = GetNextSelection(hSelectionCL);
				GetGeometry(hCentreLine, &xyStart, &xyEnd);
				}
		else {
				Close ("selection", hSelectionCL );
				Exit(%cancel, "Centre line NOT FOUND!\n");
				}
		Close ("selection", hSelectionCL );
		
		// Mirror the curves
		nn = 0 ;
		hSelectionCurves = Open("selection", "DB ID = '" + sProfileID + "' AND ( DB curvetype = 'BackCurve' OR DB curvetype = 'BackCurveLargest')");
				if (hSelectionCurves) {
				ResetSelection(hSelectionCurves);
				while ( hProfile = GetNextSelection(hSelectionCurves)) {
					// Mirror Profile about Centre Line
					// Assume that CL is horizontal
    					nVertexCount = GetVertexCount(hProfile ) ;
     					ii = 1; nn = nn + 1;
    					if (nVertexCount < 2) {
						Close ("selection", hSelectionCurves );
    						Exit(%cancel, "Can't copy selected profile!\n");
    						}
    					StartPoly ("fitted") ;
     					  while ( ii <= nVertexCount) {
				  		GetVertex (hProfile, ii, &xyPt1) ;
				   		xyPt1.y = xyStart.y - (xyPt1.y - xyStart.y);
				   		AddVertex (xyPt1) ;
				  		ii = ii + 1;
						} 
    					EndPoly();
    				
					// Set data    				
    					hNewProfile = UID("find", UID("getmax"));
    					SetDBData (hNewProfile, "curvetype", sProfileID) ;
    				
    					GetEntityData(hProfile,"layer", &hLayer);
    					SetEntityData(hNewProfile, "layer", hLayer);
    				
    					// Label the curve if required
    					GetDBData (hProfile, "Leg", &sSide);
    					if (!StringCompare(sSide, "Left&Right")){
    						SetData ("TextVertJust", 32);
    						SetData ("TextHorzJust", 2);
    						hLayerNotes = Table("find", "layer", "Notes");
    						GetVertex (hNewProfile,  GetVertexCount(hNewProfile)-nn, &xyPt1) ;
						hEnt = AddEntity("marker", "closed arrow",xyPt1, 0.25, 0.125, 90);
						SetEntityData(hEnt, "layer",hLayerNotes );
						hEnt = AddEntity("text", sSide,  xyPt1.x,  xyPt1.y+0.25);
						SetEntityData(hEnt, "layer",hLayerNotes );
    						}
    				}//While
		}//If
	}//If
else {
		Close ("selection", hSelectionCL );
		Exit(%cancel, "Polyline picked NOT a Body Profile!\n");
	}

Close ("selection", hSelectionCurves );
Close ("selection", hSelectionCL );
