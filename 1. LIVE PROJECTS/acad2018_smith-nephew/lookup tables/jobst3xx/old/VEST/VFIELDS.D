// File Name:	VFIELDS.D
// Purpose:	Set up for JOBST system for Vest with
//		Sleeves
//		
// Version:	1.01
// Date:		18.Jan.94
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
//---------------------------------------------------------------------------------------------
// Many

// Setup DataBase fields
// Body fields
   Table("add" , "field", "LtSCir", "number", 3, 1);
   Table("add" , "field", "RtSCir", "number", 3, 1);
   Table("add" , "field", "NeckCir", "number", 3, 1);
   Table("add" , "field", "SWidth", "number", 3, 1);
   Table("add" , "field", "S_Waist", "number", 3, 1);
   Table("add" , "field", "ChestCir", "number", 3, 1);
   Table("add" , "field", "WaistCir", "number", 3, 1);
   Table("add" , "field", "WaistCirUserFac", "number", 3, 1);
   Table("add" , "field", "S_EOS", "number", 3, 1);
   Table("add" , "field", "EOSCir", "number", 3, 1);
   Table("add" , "field", "EOSCirUserFac", "number", 3, 1);
   Table("add" , "field", "S_Breast", "number", 3, 1);
   Table("add" , "field", "BreastCir", "number", 3, 1);
   Table("add" , "field", "BreastCirUserFac", "number", 3, 1);
   Table("add" , "field", "NippleCir", "number", 3, 1);
   Table("add" , "field", "BraLtCup", "string");
   Table("add" , "field", "BraRtCup", "string");
   Table("add" , "field", "BraLtDisk", "string");
   Table("add" , "field", "BraRtDisk", "string");
   Table("add" , "field", "LtAxillaType", "string");
   Table("add" , "field", "RtAxillaType", "string");
   Table("add" , "field", "NeckType", "string");
   Table("add" , "field", "Closure", "string");
   Table("add" , "field", "BackNeckType", "string");
   Table("add" , "field", "NeckDimension", "string");
   Table("add" , "field", "BackNeckDim", "string");

// Body fields used in transfer of data to sleeve (Used is for both axilla)
// Otherwise these are for the LEFT axilla data transfer if axilla different heights
   Table("add" , "field", "AxillaBackNeckRad", "number", 6,3);
   Table("add" , "field", "AxillaFrontNeckRad", "number", 6,3);
   Table("add" , "field", "ShoulderToBackRaglan", "number", 6,3);

// These are for the Right axilla data transfer
   Table("add" , "field", "ABNRadRight", "number", 6,3);
   Table("add" , "field", "AFNRadRight", "number", 6,3);
   Table("add" , "field", "SBRaglanRight", "number", 6,3);
