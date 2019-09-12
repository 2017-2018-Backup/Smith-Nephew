// File Name:	WHFIELDS.D
// Purpose:	Set up for JOBST system.
//		Waist Height Data Base Fields
//		See also DRAFIX.INI
//		
// Version:	1.01
// Date:		11.Jan.94
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
//---------------------------------------------------------------------------------------------
// 

// Setup DataBase fields
// Body fields
   Table("add" , "field", "TOSCir", "number", 4, 1);
   Table("add" , "field", "TOSGivenRed", "number", 4, 2);
   Table("add" , "field", "TOSHt", "number", 4, 1);
   Table("add" , "field", "WaistCir", "number", 4, 1);
   Table("add" , "field", "WaistGivenRed", "number", 4, 1);
   Table("add" , "field", "WaistHt", "number", 4, 1);
   Table("add" , "field", "MidPointCir", "number", 4, 1);
   Table("add" , "field", "MidPointGivenRed", "number", 4, 1);
   Table("add" , "field", "MidPointHt", "number", 4, 1);
   Table("add" , "field", "LargestCir", "number", 4, 1);
   Table("add" , "field", "LargestGivenRed", "number", 4, 1);
   Table("add" , "field", "LargestHt", "number", 4, 1);
   Table("add" , "field", "LeftThighCir", "number", 4, 1);
   Table("add" , "field", "RightThighCir", "number", 4, 1);
   Table("add" , "field", "ThighGivenRed", "number", 4, 1);
   Table("add" , "field", "FoldHt", "number", 4, 1);
   Table("add" , "field", "Pressure", "string");
   Table("add" , "field", "CrotchStyle", "number", 4, 1);
   Table("add" , "field", "Body", "string");	//Multiple data field


// Leg fields
   Table("add", "field", "Leg","string");
   Table("add", "field", "TapeLengthsPt1","string");
   Table("add", "field", "TapeLengthsPt2","string");
   Table("add", "field", "TapeMMs","string");
//   Table("add", "field", "TapeMMs2","string"); 	From JOBSTNEW.D
   Table("add", "field", "Reduction","string");
//   Table("add", "field", "Reduction2","string");	From JOBSTNEW.D
   Table("add", "field", "Grams","string");  
//   Table("add", "field", "Grams2","string");  	From JOBSTNEW.D
   Table("add", "field", "MMAnkle", "number", 4, 1);
   Table("add", "field", "GramsAnkle", "number", 4, 1);
   Table("add", "field", "ReductionAnkle", "number", 4, 1);
   Table("add", "field", "AnkleTape", "string"); //Multiple data field

   Table("add", "field", "FirstTape", "number", 4, 1);
   Table("add", "field", "LastTape", "number", 4, 1);
   Table("add", "field", "ToeStyle", "string");   
   Table("add", "field", "HeelStyle", "string");  //Re-Inforcement of heel style
   Table("add", "field", "HeelType", "string");  //Heel type 
   Table("add", "field", "FootPleat1", "number", 4, 1);
   Table("add", "field", "TopLegPleat1", "number", 4, 1);
   Table("add", "field", "FootPleat2", "number", 4, 1);
   Table("add", "field", "TopLegPleat2", "number", 4, 1);
   Table("add", "field", "FootLength", "number", 4, 1);

// Point ID - Fields
//   Table("add","field","ID","string") ;   	From JOBSTNEW.D
   Table("add","field","Data","string") ;   

