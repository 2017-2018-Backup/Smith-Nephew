//DRAFIX Macro created - 10-03-2018  16:10:23
//Patient - SUSAN MCGINNIS, J31625, Head & Neck 
//by Visual Basic
HANDLE hLayer, hFacemask, hTitle, hChan, hEnt;
HEADNECK1.XY     xyTitleBoxOrigin,xyTitleOrigin, xyStart,TitleOrigin, xyTitleScale, xyOrigin;
ANGLE  aTitleAngle;
STRING sTitleName, sFileNo;
Table("add","field","ID","string");
Table("add","field","HeadNeck","string");
Table("add","field","units","string");
Table("add","field","Data","string");
Table("add","field","WorkOrder","string");
SetData("TextHorzJust",1);
SetData("TextVertJust",32);
SetData("TextHeight",0.125);
SetData("TextAspect",0.6);
SetData("TextFont",0);
GetUser ("xy","Indicate Start Point",&xyStart);
Display ("cursor","wait","Drawing");
UserSelection ("clear");
Execute ("menu","SetStyle",Table("find","style","bylayer"));
hLayer = Table("find","layer","Construct");
if ( hLayer > %zero && hLayer != 32768)Execute ("menu","SetLayer",hLayer);
hEnt = AddEntity("marker","cross",xyStart,0.125);
if (hEnt) SetDBData( hEnt,"ID","RCSJ31625OriginMarker");
if (hEnt) SetDBData( hEnt,"Data","RCS0000000000000000");
SetDBData(hEnt,"Fabric","Pow 160 Beige");
SetDBData(hEnt,"WorkOrder","20051485");
SetDBData(hEnt,"Data","RCS0000000000000000");
SetDBData(hEnt,"units","cm");
xyOrigin = xyStart;
sFileNo = "J31625";
hLayer = Table("find","layer","TemplateLeft");
if ( hLayer > %zero && hLayer != 32768)Execute ("menu","SetLayer",hLayer);
hEnt = AddEntity(
"line",
xyStart.x+ 1.46,xyStart.y+-2,
xyStart.x+ 1.46,xyStart.y+ 0
);
if (hEnt) SetDBData( hEnt,"ID","RCSJ31625VelcroBack");
if (hEnt) SetDBData( hEnt,"Data","RCS0000000000000000");
hLayer = Table("find","layer","Notes");
if ( hLayer > %zero && hLayer != 32768)Execute ("menu","SetLayer",hLayer);
hEnt = AddEntity(
"line",
xyStart.x+ .46,xyStart.y+ 0,
xyStart.x+ .46,xyStart.y+-2
);
hLayer = Table("find","layer","TemplateLeft");
if ( hLayer > %zero && hLayer != 32768)Execute ("menu","SetLayer",hLayer);
if (hEnt) SetDBData( hEnt,"ID","RCSJ31625NeckBack");
if (hEnt) SetDBData( hEnt,"Data","RCS0000000000000000");
hLayer = Table("find","layer","Notes");
if ( hLayer > %zero && hLayer != 32768)Execute ("menu","SetLayer",hLayer);
AddEntity("text","1 7/8"",xyStart.x+ .46,xyStart.y+-.25,0.06,0.1,0);
AddEntity("text"," Velcro",xyStart.x+ .56,xyStart.y+-.5,0.06,0.1,0);
hLayer = Table("find","layer","TemplateLeft");
if ( hLayer > %zero && hLayer != 32768)Execute ("menu","SetLayer",hLayer);
hEnt = AddEntity("poly","fitted"
,xyStart.x+-.285,xyStart.y+ 2
,xyStart.x+-.244285714285714,xyStart.y+ 2.003844
,xyStart.x+-.203571428571429,xyStart.y+ 2.015376
,xyStart.x+-.162857142857143,xyStart.y+ 2.034596
,xyStart.x+-.122142857142857,xyStart.y+ 2.061504
,xyStart.x+-.0814285714285715,xyStart.y+ 2.0961
,xyStart.x+-.0407142857142857,xyStart.y+ 2.138384
,xyStart.x+ 0,xyStart.y+ .76
);
if (hEnt) SetDBData( hEnt,"ID","RCSJ31625ChinStrapFront");
if (hEnt) SetDBData( hEnt,"Data","RCS0000000000000000");
hEnt = AddEntity("poly","fitted"
,xyStart.x+ 2,xyStart.y+ .76
,xyStart.x+ 2.031,xyStart.y+ .459
,xyStart.x+ 2.0775,xyStart.y+ .158
,xyStart.x+ 2.1395,xyStart.y+-.143
,xyStart.x+ 2.2325,xyStart.y+-.444
,xyStart.x+-.354165718518214,xyStart.y+-.162255690567348
,xyStart.x+-.180975093879377,xyStart.y+-.188451673573682
,xyStart.x+-.0491024582215754,xyStart.y+-.303737535469769
,xyStart.x+-1.87072579649339E-14,xyStart.y+-.471874891774247
,xyStart.x+ 0,xyStart.y+ 1.166875
,xyStart.x+ .00195643698965853,xyStart.y+ 1.34887923694899
,xyStart.x+ .00782496986808034,xyStart.y+ 1.53080229175619
,xyStart.x+-.0465,xyStart.y+ .72
,xyStart.x+-.124,xyStart.y+ .125
,xyStart.x+-.31,xyStart.y+ 0
,xyStart.x+-.465,xyStart.y+ 0
,xyStart.x+ .46,xyStart.y+ 0
);
if (hEnt) SetDBData( hEnt,"ID","RCSJ31625ChinStrapBack");
if (hEnt) SetDBData( hEnt,"Data","RCS0000000000000000");
hEnt = AddEntity(
"line",
xyStart.x+ 0,xyStart.y+ .76,
xyStart.x+ 2,xyStart.y+ .76
);
if (hEnt) SetDBData( hEnt,"ID","RCSJ31625ChinStrapTop");
if (hEnt) SetDBData( hEnt,"Data","RCS0000000000000000");
hEnt = AddEntity(
"line",
xyStart.x+ .46,xyStart.y+ 0,
xyStart.x+ 1.46,xyStart.y+ 0
);
if (hEnt) SetDBData( hEnt,"ID","RCSJ31625VelcroTop");
if (hEnt) SetDBData( hEnt,"Data","RCS0000000000000000");
hEnt = AddEntity("poly","fitted"
,xyStart.x+-.21,xyStart.y+-2
,xyStart.x+-.21,xyStart.y+-2
,xyStart.x+-.21,xyStart.y+-.375
,xyStart.x+-.240724489100581,xyStart.y+-.260334645655713
,xyStart.x+-.324665354357865,xyStart.y+-.176393780415646
,xyStart.x+-.439330708708454,xyStart.y+-.145669291338583
,xyStart.x+ .703661417493821,xyStart.y+-.145669291338583
,xyStart.x+ .286830708760137,xyStart.y+-.033979839629259
,xyStart.x+-.0183105482051829,xyStart.y+ .271161417273475
,xyStart.x+-.13,xyStart.y+ .687992125984252
,xyStart.x+-.13,xyStart.y+ 1.34399606299213
,xyStart.x+-.13,xyStart.y+ 2
);
if (hEnt) SetDBData( hEnt,"ID","RCSJ31625Chin");
if (hEnt) SetDBData( hEnt,"Data","RCS0000000000000000");
hEnt = AddEntity(
"line",
xyStart.x+-.13,xyStart.y+ 2,
xyStart.x+-.285,xyStart.y+ 2
);
if (hEnt) SetDBData( hEnt,"ID","RCSJ31625ChinStrapMouth");
if (hEnt) SetDBData( hEnt,"Data","RCS0000000000000000");
hEnt = AddEntity(
"line",
xyStart.x+-.21,xyStart.y+-2,
xyStart.x+ 1.46,xyStart.y+-2
);
if (hEnt) SetDBData( hEnt,"ID","RCSJ31625ChinStrapBottom");
if (hEnt) SetDBData( hEnt,"Data","RCS0000000000000000");
hLayer = Table("find","layer","Construct");
if ( hLayer > %zero && hLayer != 32768)Execute ("menu","SetLayer",hLayer);
hEnt = AddEntity("poly","fitted"
,xyStart.x+-.085,xyStart.y+-2
,xyStart.x+-.085,xyStart.y+-2
,xyStart.x+-.085,xyStart.y+-.375
,xyStart.x+-.1324713136318,xyStart.y+-.197834645648312
,xyStart.x+-.262165354372667,xyStart.y+-.0681406049340449
,xyStart.x+-.439330708734092,xyStart.y+-.0206692913385827
,xyStart.x+ .703661417468183,xyStart.y+-.0206692913385828
,xyStart.x+ .349330708745335,xyStart.y+ .0742733358523417
,xyStart.x+ .089942627263599,xyStart.y+ .333661417280876
,xyStart.x+-.005,xyStart.y+ .687992125984252
,xyStart.x+-.005,xyStart.y+ 1.34399606299213
,xyStart.x+-.005,xyStart.y+ 2
);
if (hEnt) SetDBData( hEnt,"ID","RCSJ31625LipStrapChinConstructLine");
if (hEnt) SetDBData( hEnt,"Data","RCS0000000000000000");
hLayer = Table("find","layer","Notes");
if ( hLayer > %zero && hLayer != 32768)Execute ("menu","SetLayer",hLayer);
AddEntity("text","Strap Length:    1/2"",xyStart.x+ .1,xyStart.y+ .26,0.06,0.1,0);
hLayer = Table("find","layer","TemplateLeft");
if ( hLayer > %zero && hLayer != 32768)Execute ("menu","SetLayer",hLayer);
hLayer = Table("find","layer","Notes");
if ( hLayer > %zero && hLayer != 32768)Execute ("menu","SetLayer",hLayer);
hLayer = Table("find","layer","Construct");
if ( hLayer > %zero && hLayer != 32768)Execute ("menu","SetLayer",hLayer);
hLayer = Table("find","layer","TemplateLeft");
if ( hLayer > %zero && hLayer != 32768)Execute ("menu","SetLayer",hLayer);
HEADNECK1.XY     xyMPD_Origin, xyMPD_Scale ;
STRING sMPD_Name;
ANGLE  aMPD_Angle;
hMPD = UID ("find",0);
if (hMPD)
  GetGeometry(hMPD, &sMPD_Name, &xyMPD_Origin, &xyMPD_Scale, &aMPD_Angle);
else
  Exit(%cancel,"Can't find > mainpatientdetails < symbol, Insert Patient Data");
if ( Symbol("find","HEADNECK")){
  Execute ("menu","SetLayer",Table("find","layer","Data"));
  hFacemask = AddEntity("symbol","HEADNECK",xyMPD_Origin);
  }
else
  Exit(%cancel, "Can't find >HEADNECK< symbol to insert\nCheck your installation, that JOBST.SLB exists!");
SetDBData(hFacemask,"HeadNeck","050050050050050050050000000000000");
SetDBData(hFacemask,"Fabric","Pow 160 Beige");
SetDBData(hFacemask,"WorkOrder","20051485");
SetDBData(hFacemask,"Data","RCS0000000000000000");
