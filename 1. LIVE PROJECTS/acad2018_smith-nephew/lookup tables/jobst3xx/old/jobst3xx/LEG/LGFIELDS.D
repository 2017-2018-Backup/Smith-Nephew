// File Name:	LGFIELDS.D
// Purpose:	Set up for JOBST system.
//		Leg \ Lower Extremity Data Base Fields
//		See also JOBSTNED.D
//		
// Version:	1.01
// Date:		19.Oct.95
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
//---------------------------------------------------------------------------------------------
// 

// Setup DataBase fields

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
   Table("add" , "field", "Pressure", "string");

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

// Leg and lower options - Fields these are used as multi data fields
// 
   Table("add","field","Anklet","string") ;   
   Table("add","field","ThighLength","string") ;   
   Table("add","field","KneeLength","string") ;   
   Table("add","field","ThighBand","string") ;   
   Table("add","field","ThighBandBK","string") ;   
   Table("add","field","KneeBand","string") ;   
