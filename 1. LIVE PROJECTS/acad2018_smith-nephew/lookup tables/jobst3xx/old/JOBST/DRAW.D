//DRAFIX Macro created - 05-03-2018  08:52:52
//Patient - STANLEY HUMPHREY, J21520
//by Visual Basic, Waist Height Leg
//type - Save Data ONLY
UserSelection ("clear");
Execute ("menu","SetColor",Table("find","color","bylayer"));
Execute ("menu","SetStyle",Table("find","style","bylayer"));
Display ("cursor","wait","Save Data ONLY");
HANDLE hMPD, hLeg;
hLeg = UID ("find",1389);
if (!hLeg) Exit(%cancel,"Can't find >waistleg< symbol to update!");
SetDBData( hLeg,"fileno","J21520");
SetDBData( hLeg,"Leg","Right");
SetDBData( hLeg,"TapeLengthsPt1","                                                            ");
SetDBData( hLeg,"TapeLengthsPt2","                     400 420 430 440                        ");
SetDBData( hLeg,"AnkleTape","-1 0 0 0 0 0 0 2");
SetDBData( hLeg,"FirstTape","21");
SetDBData( hLeg,"LastTape","24");
SetDBData( hLeg,"ToeStyle","");
SetDBData( hLeg,"FootPleat1","");
SetDBData( hLeg,"FootPleat2","");
SetDBData( hLeg,"TopLegPleat1","2.7");
SetDBData( hLeg,"TopLegPleat2","");
SetDBData( hLeg,"FootLength","");
SetDBData( hLeg,"Data","1");
