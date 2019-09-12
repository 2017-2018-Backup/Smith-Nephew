// File Name:	Select.D
// Purpose: 	To select a drawing and export Choice and
//		Modification list to Head and Neck form
// Version:	1
// Date:		25/Jul/95
// By:		Ciaran McKavanagh

HANDLE
	hSel, 		hChan, 		hDDE,	
	hEntity;

NUMBER
	nUID,		nAngle;

XY	
	xyO,		xyPt;

STRING
	sFileNo,		sPath,		sPatientName,
	sUnits,		sSex,		sAge,	sPathJOBST ,
	sDiagnosis,	sWorkOrder,	sFabric,
	sData,		sHeadNeck,	sUID,
	sSymbolName,	sDataString;

// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "JOBST", "DRAFIX.INI") ;


hEntity = GetEntity("Select part of a Design", "all");

if (hEntity)
	{
	GetDBData(hEntity, "Data", &sDataString);
		
	Execute ("application", sPathJOBST +"\\HeadNeck\\HeadNeck", "normal");
	hDDE = Open ( "dde",  "HeadNeck", "Form1");
	
	Poke(hDDE, "txtData", sDataString);
	GetData("DrawingFile", &sPath);
	Poke(hDDE, "txtDraw", sPath);	

	//get data from drawing and poke to vb
	// If "titlebox" already exists use data from it

   	hChan = Open ("selection", "DB SymbolName = 'mainpatientdetails'");
   	ResetSelection(hChan);
   	hSel = GetNextSelection(hChan);
  
  	if (hSel) {
		GetGeometry(hSel, &sSymbolName, &xyO, &xyPt, &nAngle);
		GetDBData( hSel, "patient", &sPatientName);
		Poke(hDDE, "txtPatientName", sPatientName);
   		GetDBData( hSel, "fileno", &sFileNo);
		Poke(hDDE, "txtFileNo", sFileNo);
		GetDBData( hSel, "units", &sUnits);
		Poke(hDDE, "txtUnits", sUnits);
		GetDBData( hSel, "sex", &sSex);
		Poke(hDDE, "txtSex", sSex);
		GetDBData( hSel, "diagnosis", &sDiagnosis);
		Poke(hDDE, "txtDiagnosis", sDiagnosis);
		GetDBData( hSel, "age", &sAge);
		Poke(hDDE, "txtAge", sAge);
		GetDBData( hSel, "WorkOrder", &sWorkOrder); Poke(hDDE, "txtWorkOrder", sWorkOrder);
		Poke ( hDDE, "txtUidMPD", MakeString("long",UID("get",hSel)));
		}

	Close("selection",hChan);

  	hChan = Open ("selection", "DB SymbolName = 'headneck'");
  	ResetSelection(hChan);
   	hSel = GetNextSelection(hChan);
  
 	if (hSel) {
		GetGeometry(hSel, &sSymbolName, &xyO, &xyPt, &nAngle);
		GetDBData( hSel, "HeadNeck", &sHeadNeck);
		Poke(hDDE, "txtMeasurements", sHeadNeck);
		GetDBData( hSel, "Fabric", &sFabric);
		Poke(hDDE, "txtFabric", sFabric);
		Poke ( hDDE, "txtUidHN", MakeString("long",UID("get",hSel )));
		}
	Close("selection",hChan);

	if (!hDDE) Display ( "message", "error", "Can't start a DDE Link - Retry");
	Close ( "dde", hDDE);
	}
