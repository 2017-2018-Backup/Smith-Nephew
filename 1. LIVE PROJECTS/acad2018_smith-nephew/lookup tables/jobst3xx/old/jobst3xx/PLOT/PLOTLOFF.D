// File Name:	PLOTLOFF.D
// Purpose:	To Switch off layers for plotting
//		
// Version:	1.01
// Date:		14.Nov.93
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
//---------------------------------------------------------------------------------------------
//

// Local Variables 
HANDLE
	hChan,
	hTitle,
	hMasked,
	hLayer;

STRING
	sFileNo ;

NUMBER
	ii;

// NB.       Explicit conditional statement  "hLayer > %zero && hLayer != 32768"
//	Because it don't work otherwise.

   UserSelection("clear");

// Get FileNo Details
//
   hChan = Open ("selection", "DB SymbolName = 'mainpatientdetails'");
   ResetSelection(hChan);
   hTitle = GetNextSelection(hChan);
   if (hTitle) {
   	GetDBData( hTitle, "fileno", &sFileNo);
  	Close("selection",hChan);
  	}
        else 
	Exit(%cancel, "Patient details cannot be found\nPlease ensure that TITLE BOX has been used\nThen try again\n");
	

// Make 1 the current layer.
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;

// Switch off the following layers
//	1. Construct
//	2. titlebox
//	3. Data
//	4. Pattern
//	5. joblog

   hMasked = Table("find","state", "masked") ;

   hLayer = Table("find", "layer", "Construct");
   if ( hLayer > %zero && hLayer != 32768)
	SetTableData( "layer", hLayer, "state", hMasked) ;

   hLayer = Table("find", "layer", "titlebox");
   if (hLayer > %zero && hLayer != 32768)
	SetTableData( "layer", hLayer, "state", hMasked) ;

   hLayer = Table("find", "layer", "Data");
   if ( hLayer > %zero && hLayer != 32768 )
	SetTableData( "layer", hLayer, "state", hMasked) ;

   hLayer = Table("find", "layer", "Pattern");
   if ( hLayer > %zero && hLayer != 32768)
	SetTableData( "layer", hLayer, "state", hMasked) ;

   hLayer = Table("find", "layer", "joblog");
   if ( hLayer > %zero && hLayer != 32768 )
	SetTableData( "layer", hLayer, "state", hMasked) ;

   hLayer = Table("find", "layer", "OLD");
   if ( hLayer > %zero && hLayer != 32768 )
	SetTableData( "layer", hLayer, "state", hMasked) ;

// Swich off sleeve layers left and right
   ii = 0 ;
   while ( ii <= 18) {
	// Left sleeve
	hLayer = Table("find", "layer", sFileNo + "L" + MakeString("long", ii)) ;
  	if ( hLayer > %zero && hLayer != 32768  )
		SetTableData( "layer", hLayer, "state", hMasked) ;
	// Right sleeve
 	hLayer = Table("find", "layer", sFileNo + "R" + MakeString("long", ii)) ;
  	if ( hLayer > %zero && hLayer != 32768 )
		SetTableData( "layer", hLayer, "state", hMasked) ;
	ii = ii + 1;
	}
   hLayer = Table("find", "layer", sFileNo + "L100");
   if ( hLayer > %zero && hLayer != 32768 )
	SetTableData( "layer", hLayer, "state", hMasked) ;
   hLayer = Table("find", "layer", sFileNo + "R100");
   if ( hLayer > %zero && hLayer != 32768 )
	SetTableData( "layer", hLayer, "state", hMasked) ;


// Refresh screen
   Execute("menu", "ViewRedraw") ;