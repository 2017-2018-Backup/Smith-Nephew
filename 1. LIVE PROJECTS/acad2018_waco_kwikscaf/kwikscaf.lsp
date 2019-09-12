;;;	(c) Waco Kwikform Limited
;;;	ACN 002 835 36
;;;	P.O. Box 15 Rydalmere NSW 2116
;;;
;;;	All rights reserved. No part of this work covered by copyright may be
;;;	reproduced or copied in any form or by any means (graphic, electronic
;;;	or mechanical, including photocopying, recording, recording taping or
;;;	information retrieval system) without the written permission of Waco
;;;	Kwikform Limited.



;;;=====================================================================|
;;;     Function: KwikScaf.LSP						|
;;;=====================================================================|
;;;  Description: Basic Lisp functions that are used in KwikScaf.	|
;;;=====================================================================|



;;;---------------------------------------------------------------------|
;;;     Function: KSARX							|
;;;---------------------------------------------------------------------|
;;;  Description: Loads KwikScaf.arx or returns an error if it is.	|
;;;---------------------------------------------------------------------|
;;; This function loads the KwikScaf.arx. If it is loaded already then	|
;;; an error message is displayed stating that the ARX is loaded.	|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;	<none>								|
;;;---------------------------------------------------------------------|

(DEFUN C:KSARX ()			;Defines the function named KSARX
  (CEC)					;Calls the "cec" function to check the command echo
  (SETVAR "cmdecho" 0)			;sets the command echo to off
 
  (SETVAR "cmdecho" CMDE)		;Sets the command echo back to it previous state
)					;Closes the DEFUN



;;;---------------------------------------------------------------------|
;;;     Sub-routine: CEC						|
;;;---------------------------------------------------------------------|
;;;  Description: Command echo checker and setter.			|
;;;---------------------------------------------------------------------|
;;; Looks at the command line echo system variable and set it to cmde	|
;;; from the other application to use.					|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;	cmde			;;the current setting of the command	|
;;;				;;echo is taken from "cmdecho" system	|
;;;				;;variable. This variable is global	|
;;;---------------------------------------------------------------------|

(DEFUN CEC ()				;Defines the function named cec
  (SETQ CMDE (GETVAR "cmdecho"))	;Set cmde to the current Value of "cmdecho" system variable
)					;Closes DEFUN




;;;---------------------------------------------------------------------|
;;;     Function: KSM							|
;;;---------------------------------------------------------------------|
;;;  Description: Checks if menu KwikScaf is load.			|
;;;---------------------------------------------------------------------|
;;; Checks to see if the KwikScaf menu is loaded. If it is not then the	|
;;; the function KSM is called.						|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;	KSM			;;looks for the menu group name KwikScaf|
;;;				;;and set it to this variable.		|
;;;---------------------------------------------------------------------|

;(DEFUN C:KSSM (/ KSM)			;Defines the function KSSM and set the variable KSM to local
 ; (CEC)					;Calls the CEC function
;  (SETVAR "cmdecho" 0)			;Sets the command echo to off
;  (SETQ KSM (MENUGROUP "KwikScaf"))	;Looks for the menu group KwikScaf and set the value to the variable KSM
;  (IF (= nil KSM)			;IF KSM is nil ...
;    (KSML)				;THEN the function KSML
;    (PRINC "\nKwikScaf menu is already loaded.")
;					;ELSE print the message that the menu is loaded
;  )					;Close IF
;  (SETVAR "cmdecho" CMDE)		;Set the command echo back to it original setting
;)					;Close DEFUN



;;;---------------------------------------------------------------------|
;;;     Sub-routine: KSML						|
;;;---------------------------------------------------------------------|
;;;  Description: Loads KwikScaf menu file and set it to be pop10.	|
;;;---------------------------------------------------------------------|
;;; Loads the KwikScaf menu from the KwikScaf.MN? file and then set it	|
;;; to be pop menu 10. A message is displayed once the menu has been	|
;;; loaded.								|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;	KSM			;;looks for the menu group name KwikScaf|
;;;				;;and set it to this variable.		|
;;;---------------------------------------------------------------------|

;(DEFUN KSML (/)				;Defines the function and reset all local variables
;  (CEC)					;Calls the CEC function
;  (SETVAR "cmdecho" 0)			;Sets the command echo to off
;  (COMMAND "menuload" "KwikScaf")	;Issues the AutoCAD command "menuload" and prompts it to load the KwikScaf menu
;  (MENUCMD "P10=+KwikScaf.pop1")	;Places the KwikScaf menu at position 10
;  (PRINC "\nKwikScaf menu has now been loaded.")
;					;Returns a string to the command line telling the use that the menu has been loaded
;  (SETVAR "cmdecho" CMDE)		;Set the command echo back to it original setting
;)					;Close DEFUN


;;;---------------------------------------------------------------------|
;;;     Function: PSMS							|
;;;---------------------------------------------------------------------|
;;;  Description: Check for status of the current space.		|
;;;---------------------------------------------------------------------|
;;; Looks at the current space that the drawing is in and change to 	|
;;; model space if in paper space.					|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;	SPACE			;;The current value of the "tielmode"	|
;;;				;;system variable.			|
;;;---------------------------------------------------------------------|

(DEFUN C:PSMS (/ SPACE)			;Defines the function and reset all local variables
  (CEC)					;Calls the CEC function
  (SETVAR "cmdecho" 0)			;Sets the command echo to off
  (SETQ SPACE (GETVAR "tilemode"))	;Looks at the system variable "tilemode" and sets it value to the variable "SPACE"
  (IF (= SPACE 0)			;IF the value of space is equal to 0 ...
    (SETVAR "tilemode" 1)		;THEN set the value of the system variable of "tilemode" to 0
  )					;Close IF
  (SETVAR "cmdecho" CMDE)		;Set the command echo back to it original setting
)					;Close DEFUN



;;;---------------------------------------------------------------------|
;;;     Function: KSZM							|
;;;---------------------------------------------------------------------|
;;;  Description: Zooms to 0.001 of the current space.			|
;;;---------------------------------------------------------------------|
;;; Zooms the current space using the zoom scale command so as not to	|
;;; have to manually zoom or RT zoom.					|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;	NONE								|
;;;---------------------------------------------------------------------|

(DEFUN C:KSZM (/)			;Defines the function and reset all local variables
  (CEC)					;Calls the CEC function
  (SETVAR "cmdecho" 0)			;Sets the command echo to off
  (PRINC "\nZooming to 0.01 of limits")	;Tells the user that the screen will be zoomed to 0.01
  (COMMAND "zoom" "0.01")		;Zooms the screen to 0.01
  (PRINC "\nPaning to positive screen.");Tells the user that the screen will be paned
  (COMMAND "-pan" "50000,28000" "0,0")	;Pans the screen
  (SETVAR "cmdecho" CMDE)		;Set the command echo back to it original setting
)					;Close DEFUN



;;;---------------------------------------------------------------------|
;;;     Function: PSKST							|
;;;---------------------------------------------------------------------|
;;;  Description: Sets "pickstyle" system variable.			|
;;;---------------------------------------------------------------------|
;;; Sets the "pickstyle" system variable to 1 so that the groups will	|
;;; pick selected when working with KwikScaf.				|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;	PSS			;;Set to the current value of the	|
;;;				;;"pickstyle" system variable		|
;;;---------------------------------------------------------------------|

(DEFUN PKST (/ PSS)			;Defines the function and reset all local variables
  (CEC)					;Calls the CEC function
  (SETVAR "cmdecho" 0)			;Sets the command echo to off
  (SETQ PSS (GETVAR "pickstyle"))	;Sets PSS to the current value of the "pickstyle" system variable
  (IF (/= PSS 1)			;IF PSS is less than or equal to 1 ...
    (SETVAR "pickstyle" 1)		;THEN the "pcikstely" system variable is set to 1
  )					;Close IF
  (SETVAR "cmdecho" CMDE)		;Set the command echo back to it original setting
)					;Close DEFUN



;;;---------------------------------------------------------------------|
;;;     Function: KSLT							|
;;;---------------------------------------------------------------------|
;;;  Description: Checks for KwikScaf Line Types & loads if needed	|
;;;---------------------------------------------------------------------|
;;; Looks to see if the KwikScaf line types are loaded in the current	|
;;; drawing. If the line types are not loaded they will be loaded.	|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;	COUNTER			;;Sets a loop counter			|
;;;	LTLS			;;List of line style names to check for	|
;;;	LTS			;;Name of the line style to check for	|
;;;	KSLT			;;Value of line style from table	|
;;;	LSTR			;;String to tell use outcome of attempt	|
;;;---------------------------------------------------------------------|

(DEFUN c:KSLT (/ COUNTER LTLS LTS KSLT LSTR)
					;Defines the function and reset all local variables
  (CEC)					;Calls the CEC function
  (SETVAR "cmdecho" 0)			;Sets the command echo to off
  (SETQ	COUNTER	0			;Sets COUNTER to the value of 0
	LTLS	'("KwikScaf_Line"
		  "KwikScaf_Dashed"
		  "KwikScaf_DOTD"
		  "KwikScaf_DOTC"
		 )			;Sets LTLS to the list of line style names to look for
  )					;Close DEFUN
  (WHILE (<= COUNTER 3)			;Set a WHILE loop the runs as long and COUNTER is less than or equal to 3
    (SETQ LTS  (NTH COUNTER LTLS)	;Sets LTS to the NTH COUNTER of the LTLS list
	  KSLT (TBLSEARCH "Ltype" LTS)	;Sets KSLT to value returned for the table search of the "Ltype" for the value of LTS
    )					;Close SETQ
    (IF	(= KSLT nil)			;IF KSLT is equal to nil ...
      (PROGN				;Start PROGN for THEN
	(SETQ
	  LSTR (STRCAT LTS " line type not loded.\nLoading " LTS ".")
	)				;Sets LSTR to a string that tells the user that the value of LTS is not currently loaded
	(PRINC LSTR)			;Prints the string of LSTR
	(COMMAND "linetype" "Load" LTS "kwikscaf.lin" "")
					;Loads the line type LTS
      )					;Close PROGN for THEN
      (PROGN				;Starts PROGN for ELSE
	(SETQ LSTR (STRCAT "\n" LTS " line type is loaded."))
					;Sets the LSTR to a sting that tells the user that the LTS line type is already loaded
	(PRINC LSTR)			;Prints the string of LSTR
      )					;Close PROGN for ELSE
    )					;Close IF
    (SETQ COUNTER (+ COUNTER 1))	;Increments the counter by 1
  )					;Close WHILE
  (SETVAR "cmdecho" CMDE)		;Set the command echo back to it original setting
)					;Close DEFUN



;;;---------------------------------------------------------------------|
;;;     Function: DLGRESET						|
;;;---------------------------------------------------------------------|
;;;  Description: Resets the KwikScaf Dialogue boxes locations		|
;;;---------------------------------------------------------------------|
;;; Resets the registry keys for the locations of the KwikScaf dialogue |
;;; boxed. All the keys are deleted so that when the user will moves the|
;;; dialogue box the keys will be recreated.				|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;	RegKeyLst		;;A list of registry keys to be removed	|
;;;	LoopCount		;;Length of the list used for looping	|
;;;	CurRegKey		;;The current value for the RegKeyLst	|
;;;---------------------------------------------------------------------|

(DEFUN c:DLGRESET (/ RegKeyLst LoopCount CurRegKey)
					;Defines the function and reset all local variables
  (CEC)					;Calls the CEC function
  (SETVAR "cmdecho" 0)			;Sets the command echo to off
  (SETQ	RegKeyLst '("AssignStageDialog_X"
		     "AssignStageDialog_Y"
		     "AutobuildDialog_X"
		     "AutobuildDialog_Y"
		     "AutoBuildPropDialog_X"
		     "AutoBuildPropDialog_Y"
		     "BOMExtra_X"
		     "BOMExtra_Y"
		     "BOMSelectionDialog_X"
		     "BOMSelectionDialog_Y"
		     "BOMSummaryDialog_X"
		     "BOMSummaryDialog_Y"
		     "BayDetailsDialog_X"
		     "BayDetailsDialog_Y"
		     "BayPropertiesDialog_X"
		     "BayPropertiesDialog_Y"
		     "ChangeRLDlg_X"
		     "ChangeRLDlg_Y"
		     "CKwikscafProgressDlg_X"
		     "CKwikscafProgressDlg_Y"
		     "CListPrintPage1_X"
		     "CListPrintPage1_Y"
		     "CListPrintPage2_X"
		     "CListPrintPage2_Y"
		     "CListPrintSetup_X"
		     "CListPrintSetup_Y"
		     "CPrintStatus_X"
		     "CPrintStatus_Y"
		     "EditStandardsDlg_X"
		     "EditStandardsDlg_Y"
		     "HelpAboutDlg_X"
		     "HelpAboutDlg_Y"
		     "JobDetailsDialog_X"
		     "JobDetailsDialog_Y"
		     "LiftDetailsDialog_X"
		     "LiftDetailsDialog_Y"
		     "RunPropDialog_X"
		     "RunPropDialog_Y"
		     "SetStagesDlg_X"
		     "SetStagesDlg_Y"
		     "StandardHeightPickerDialog_X"
		     "StandardHeightPickerDialog_Y"
		    )			;Sets RegKeyLst to a list of Registry key names used by KwikScaf dialogue boxes
	LoopCount (LENGTH RegKeyLst)	;Sets a loop counter to the number of keys in the RegKeyLst list
  )					;Close SETQ
  (WHILE (> LoopCount 0)		;WHILE LoopCount is great than 0 ...
    (SETQ LoopCount (1- LoopCount)	;Sets the loop counter to 1 less than what it was
	  CurRegKey (NTH LoopCount RegKeyLst)
					;Sets the CurRegKey to the NTH LoopCount of the RegKeyLst
    )					;Close SETQ
    (VL-REGISTRY-DELETE
      "HKEY_CURRENT_USER\\Software\\Autodesk\\AutoCAD\\R15.0\\ACAD-1:409\\Meccano"
      CurRegKey
    )					;Visual LISP command to delete a registry key matching the value of CurRegKey
  )					;Close WHILE
  (SETVAR "cmdecho" CMDE)		;Set the command echo back to it original setting
)					;Close DEFUN


;;;---------------------------------------------------------------------|
;;;     Function: TitleBlockEdit					|
;;;---------------------------------------------------------------------|
;;;  Description: This function get all object that are in the		|
;;;               crossing windows of -1,-1 to 1,1 and then searches	|
;;;               for any valid KwikScaf Title blocks.			|
;;;---------------------------------------------------------------------|
;;;  If there is nothing in the crossing windows then the function	|
;;;  alerts the user that there is nothing to check. If there are	|
;;;  valid KwikScaf blocks in the selection then the user is alerted.	|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;   CroSS                 ;; the selection set from -1,-1 to 1,1	|
;;;   CrossNo               ;; the number of objects in the CroSS	|
;;;			    ;; selection minus 1 used for looping	|
;;;   NameAss               ;; the name of the block that is the	|
;;;			    ;; current SelName object			|
;;;   NameCount             ;; loop counter for the number of times	|
;;;			    ;; to check each object in the set		|
;;;   NameList              ;; a list of names to check for		|
;;;   NameTest              ;; one of the names for NameList		|
;;;   SelName               ;; the name of the current object		|
;;;   KSBlock               ;; check variable to say if there is a	|
;;;			    ;; KwikScaf block in the current layout	|
;;;---------------------------------------------------------------------|


(DEFUN c:TitleBlockEdit	(/	    CroSS      CrossNo	  NameAss
			 NameCount  NameList   NameTest	  KSBlock
			)		;defin application and variables
  (CEC)					;Calls the CEC fuinction
  (SETVAR "cmdecho" 0)			;Sets the command echo to off
  (SETQ CroSS (SSGET "C" '(-1 -1) '(1 1)))
					;sets the CroSS variable to a list of objects at 0,0,0
  (IF (NULL CroSS)			;checks the value of the CroSS variable if it is Null
    (ALERT "There is nothing at 0,0,0 to check!")
					;then CroSS variable is less Null give an alert
    (PROGN				;else ...
      (SETQ CrossNo   (- (SSLENGTH CroSS) 1)
					;sets CrossNo to 1 less than the number of objects in the list for CroSS
	    NameList  '("KwikScaf A0 Drawing Sheet"
			"KwikScaf A1 Drawing Sheet"
			"KwikScaf A2 Drawing Sheet"
			"KwikScaf A2 Drawing Sheet No Ledgen"
			"KwikScaf A3 Drawing Sheet No Ledgen"
		       )		;sets NameList to a list of valid title block to check
	    NameCount (- (LENGTH NameList) 1)
					;sets NameCount to a loop counter of the number of names in NameList minus 1
      )					;end setq
      (WHILE (>= CrossNo 0)		;set the condition that while CrossNo is greater then or equal to 0 ....
	(SETQ SelName (SSNAME CroSS CrossNo))
					;sets SelName to the object name of the name from the CrossNo (nth) of CroSS
	(WHILE (>= NameCount 0)		;sets the condition that while NameCounter id greater then or equal to 0 ...
	  (SETQ	NameAss	 (CDR (ASSOC 2 (ENTGET SelName)))
					;sets NameAss to the name of the block of SelName
		NameTest (NTH NameCount NameList)
					;sets NameTest to the NameCount (nth) of the NameList
	  )				;end setq
	  (IF (= NameAss NameTest)	;checks if NameAss equals NameTest
	    (PROGN			;if NameAss does equal NameTest then ...
	      (COMMAND "ddatte" SelName);issues the edit attribute command on the object SelName
	      (SETQ KSBlock 1)		;and then sets the KSBlock to 1 to say that there has been a KSBlock edited
	    )				;end progn for NameAss does equal NameTest
	  )				;end if NameAss equals NameTest
	  (SETQ NameCount (- NameCount 1)) ;sets NameCount to one less
	)				;end while NameCount >= 0
	(SETQ CrossNo	(- CrossNo 1)	;sets CrossNo to one less
	      NameCount	4		;resets NameCount to 4
	)				;ends setq
      )					;end while CrossNo >= 0
      (IF (NULL KSBlock)		;checks if KSBlock is NUll
	(ALERT
	  "There are no valid KwikScaf title blocks in this layout!"
	)				;if KSBlock is null then alert that there were no blocks found
      )					;end if KSBlock Null
    )					;end progn for the if CroSS NULL else section
  )					;end if CroSS Null
  (SETVAR "cmdecho" CMDE)		;Set the command echo back to it original setting
)					;Close DEFUN



;;;---------------------------------------------------------------------|
;;;	Function: AttPushProp						|
;;;---------------------------------------------------------------------|
;;;  Description: AttPushProp is used for forcing the values from	|
;;;		  the Drawing Properties to the title block.		|
;;;---------------------------------------------------------------------|
;;;  This command will only work for the KwikScaf title blocks and the	|
;;;  requires that the user pick the any were on the title block objects|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;   AttName		    ;; Name of the attribute to fill		|
;;;   NoPTLst               ;; Number in the point list to check for	|
;;;   Loop1Chk              ;; Loop counter for number of items in the	|
;;;			    ;; point list for the Drawing Properties	|
;;;   DefChk                ;; Current point to check for		|
;;;   AttDef                ;; Value of the current attribute checked	|
;;;   AttName               ;; Name of attribute to check		|
;;;   AttDataEx             ;; Date extracted for the attribute		|
;;;   AttDataNew            ;; Constructed list to replace the existing	|
;;;			    ;; attributes with				|
;;;   NamePtList            ;; Named point list of Drawing Properties	|
;;;   SubEnt		    ;; Name of the object to be substituted with|
;;;			    ;; with the constructed list AttDataEx	|
;;;---------------------------------------------------------------------|
;;;	Break down of point list					|
;;;	(2	.	Project_Title)					|
;;;	(4	.	Draftsman)					|
;;;	(6	.	Project_Description)				|
;;;	(301	.	DWG_NO)						|
;;;	(304	.	Client_Name)					|
;;;	(305	.	Company)					|
;;;	(306	.	Address_Street)					|
;;;	(307	.	City)						|
;;;	(308	.	Phone_No)					|
;;;	(309	.	Fax_No)						|
;;;---------------------------------------------------------------------|

(DEFUN C:AttPushProp (/		  AttName     NoPTLst	  Loop1Chk
		      DefChk	  AttDef      AttName	  AttDataEx
		      AttDataNew  NamePtList  SubEnt	  WarnStr
		      DwgPropData
		     )			;defin application and variables
  (CEC)					;Calls the CEC fuinction
  (SETVAR "cmdecho" 0)			;Sets the command echo to off
  (SETQ	NamePtList
	 (REVERSE			;set NamePtList to an associate list
	   (LIST			;start of list to build
	     (CONS 2 "PROJECT_TITLE:")
	     (CONS 4 "DRAFTSMAN")
	     (CONS 6 "PROJECT_DESCRIPTION")
	     (CONS 301 "DWG_NO")
	     (CONS 304 "CLIENT_CONTACT")
	     (CONS 305 "COMPANY")
	     (CONS 306 "ADDRESS_STREET")
	     (CONS 307 "CITY")
	     (CONS 308 "PHONE_NO.")
	     (CONS 309 "FAX_NO.")
	   )				;Close LIST
	 )
  )					;Close SETQ
  (DwgProps)				;Call DwgProps
  (SETQ	AttName	 (ENTNEXT (CDR (BlockGet)))
					;Sets AttName to the next object in BlockGet Selection
	NoPTLst	 (1- (LENGTH NamePtList))
					;Sets NoPTLst to the length of NamePtList less 1
	Loop1Chk 1			;Sets LoopChk to 1
	DefChk	 (NTH NoPTLst NamePtList)
					;Sets DefChk to the NTH NoPTLst of NamPtList
  )					;Close SETQ
  (WHILE (>= NoPTLst 0)			;WHILE NoPTLst is greater than or equal to 0 ...
    (WHILE (= Loop1Chk 1)		;WHILE Loop1Chk is equal to 1 ...
      (SETQ AttDef (ASSOC 2 (ENTGET AttName)))
					;Sets ArrDef to the ASSOC of the AttName object
      (IF (EQUAL (CDR AttDef) (CDR DefChk))
					;IF the value of AttDef equals the vales of DefChk ...
	(SETQ Loop1Chk 0)		;THEN set the Loop1Chk to 0
	(SETQ AttName (ENTNEXT AttName));Set the AttName to the next object
      )					;Close IF
      (IF (= Attname nil)		;IF Attanem is equal to nil ...
	(SETQ Loop1Chk 0)		;THEN set the Loop1Chk to 0
      )					;Close IF
    )					;Close WHILE
    (SETQ Loop1Chk 1)			;Set Loop1Chk to 1
    (IF	(EQUAL (CDR AttDef) (CDR DefChk))
					;IF the value of AttDef equals the value of DefChk
      (PROGN				;Start then PROGN
	(SETQ AttDataEx (ASSOC 1 (ENTGET AttName)))
					;Sets AttDataEx to the 1st Assoc of the AttName Object
	(SETQ AttDataNew
	       (CONS (CAR AttDataEx)
		     (CDR (ASSOC (CAR defchk) DwgPropData))
					;Sets the AttDataNew to the constructed value of AttDataEx, DefChk and DWGPropData
	       )			;Close CONS
	)				;Close SETQ
	(IF (NOT (= nil (VL-STRING-SEARCH "\r\n" (CDR AttDataNew))))
					;IF sting of AttDataNew has "\r\n" codes
	  (CRSplit)			;Then run sub-routine CRSplit

;;; Folloing line commmented out as was only used for checking
;;;	  (WarnSplit)

	  (ModAtt)			;Else run sub-routine ModAtt
	)				;Close IF
      )					;Close PROGN
      (PROGN				;Start PROGN
	(PRINC "\nLook Again. Could not find ")
					;Tells the user that it could not find the values ..
	(PRINC DefChk)			;for DefChk
      )					;Close else PROGN
    )					;Close IF

;;;    Following lines commented out as they were only used for testing
;;;    (SETQ NowStr (STRCAT "\n"(CDR (ASSOC 2 AttDataNew))" - "(CDR (ASSOC 1 AttDataNew))))
;;;    (PRINC NowStr)

    (SETQ AttName (ENTNEXT (CDR SubEnt));Sets AttName to the next object in the SubEnt selection
	  NoPTLst (1- NoPTLst)		;Sets the NoPTLst to 1 less
    )					;Close SETQ
    (IF	(>= NoPTLst 0)			;IF NoPTLst is greater than or equal to 0 ...
      (SETQ DefChk (NTH NoPTLst NamePtList))
					;THEN set DefChk to the NTH NoPTLst of the NamePtList
    )					;Close IF
  )					;Close WHILE
  (ENTUPD (CDR (ASSOC -1 AttDataNew)))	;updates the block
  (SETVAR "cmdecho" CMDE)		;Set the command echo back to it original setting
)					;Close DEFUN



;;;---------------------------------------------------------------------|
;;;	Sub-routine: DwgProps						|
;;;---------------------------------------------------------------------|
;;;  Description: Used to create a list of all the values in the Drawing|
;;;		  properties.						|
;;;---------------------------------------------------------------------|
;;;  Looks at the named dictionary for the drawing properties and	|
;;;  creates a list for of its values.					|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;   dict		    ;; Dictionary to search for			|
;;;   title	            ;; Tittle value taken from the assco 2 of	|
;;;			    ;; the dictionary				|
;;;   author	            ;; Author value taken from the assco 4 of	|
;;;			    ;; the dictionary				|
;;;   comments	            ;; Comments value taken from the assco 6 of	|
;;;			    ;; the dictionary				|
;;;   AssocNo	            ;; The number to associate with the values	|
;;;   MarketEnqNo           ;; MarketEnqNo value taken from the assco	|
;;;			    ;; AsscoNo of the dictionary		|
;;;   ClientName            ;; ClientName value taken from the assco	|
;;;			    ;; AsscoNo of the dictionary		|
;;;   ClientPhone           ;; ClientPhone value taken from the assco	|
;;;			    ;; AsscoNo of the dictionary		|
;;;   ClientZip             ;; ClientZip value taken from the assco	|
;;;			    ;; AsscoNo of the dictionary		|
;;;   ClientFax             ;; ClientFax value taken from the assco	|
;;;			    ;; AsscoNo of the dictionary		|
;;;   ClientCompany         ;; ClientCompany value taken from the assco	|
;;;			    ;; AsscoNo of the dictionary		|
;;;---------------------------------------------------------------------|

(DEFUN DwgProps	(/	       dict	     title
		 author	       comments	     AssocNo
		 MarketEnqNo   ClientName    ClientPhone
		 ClientZip     ClientFax     ClientAddress
		 ClientCompany
		)			;defin application and variables
  (SETQ dict (MEMBER '(3 . "DWGPROPS") (ENTGET (NAMEDOBJDICT))))
					;Looks for the dwgprops dictionary
  (IF (/= dict nil)			;IF the value of dict is nil...
    (PROGN				;Start PROGN
      (SETQ dict (ENTGET (CDR (NTH 1 dict))))
					;Set dict to the nth value
      (SETQ title	  (ASSOC 2 dict)
	    author	  (ASSOC 4 dict)
	    comments	  (ASSOC 6 dict)
	    AssocNo	  301		;set AssocNo to x
	    MarketEnqNo	  (GetCustomProp dict AssocNo)
					;calls GetCustomProp and feeds it
					;AssocNo and  dict
	    AssocNo	  304
	    ClientName	  (GetCustomProp dict AssocNo)
	    AssocNo	  305
	    ClientCompany (GetCustomProp dict AssocNo)
	    AssocNo	  306
	    ClientAddress (GetCustomProp dict AssocNo)
	    AssocNo	  307
	    ClientZip	  (GetCustomProp dict AssocNo)
	    AssocNo	  308
	    ClientPhone	  (GetCustomProp dict AssocNo)
	    AssocNo	  309
	    ClientFax	  (GetCustomProp dict AssocNo)
      )					;Close SETQ
    )					;Close PROGN
  )					;Close IF


  (SETQ	DwgPropData
	 (LIST title	      author	     comments
	       MarketEnqNo    ClientName     ClientCompany
	       ClientAddress  ClientZip	     ClientPhone
	       ClientFax		;creates a list
	      )				;Close LIST
  )					;Close SETQ
  DwgPropData				;passes the value back
)					;Close DEFUN



;;;---------------------------------------------------------------------|
;;;	Sub-routine: GetCustomProp					|
;;;---------------------------------------------------------------------|
;;;  Description: Used to extract the custom properties form the DWG	|
;;;		  Props							|
;;;---------------------------------------------------------------------|
;;;  Sub-routine to extract the vales for the custom properties in the	|
;;;  AutoCAD Drawing properties.					|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;   dict		    ;; Dictionary passed by parent		|
;;;   AssocNo	            ;; Assoc number passed by parent		|
;;;   Value	            ;; Value that is passed back to the parent	|
;;;   Index	            ;; Character to look for in the string	|
;;;---------------------------------------------------------------------|

(DEFUN GetCustomProp (dict AssocNo / Value Index)
  (SETQ Value (CDR (ASSOC AssocNo dict)))
					;Gets the value of the AssocNo from the
					;dictionary
  (IF (EQUAL Value "=")			;set the if for the value being "='
    (SETQ Value nil)			;then sets Values to nil
    (PROGN				;Else starts a Progn
      (SETQ Index (VL-STRING-SEARCH "=" Value)
					;searches for the "=" character in the
					;Value and set it place to index
	    Value (CONS	AssocNo		;sets Value to a constructed list off
			(SUBSTR Value (+ Index 2))
					;the AssocNo . everything after the "="
					;in Value
		  )			;Close CONS
      )					;Close SETQ

;;;      Following IF Commented out as it was only used for checking.
;;;      (IF (NOT (= nil (VL-STRING-SEARCH "\r\n" (CDR Value))))	;IF the string of Value had \r\n codes in it ...
;;;	       (ALERT (STRCAT (CDR Value) " has CR & NL escape codes in it!"))	;THEN show an alert box showing the vales with the sting
;;;	       )	;Close IF

    )					;Close PROGN
  )					;Close IF
  Value					;passes the value back
)					;Close DEFUN



;;;---------------------------------------------------------------------|
;;;	Sub-routine: BlockGet						|
;;;---------------------------------------------------------------------|
;;;  Description: Used to find the block with in a selection		|
;;;---------------------------------------------------------------------|
;;;  Sub-routine to look for a block that may be in the section set	|
;;;  that has been passed to the routine.				|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;   PickSS		    ;; the Selection set by parent		|
;;;   AssocNo	            ;; Assoc number passed by parent		|
;;;   Value	            ;; Value that is passed back to the parent	|
;;;   Index	            ;; Character to look for in the string	|
;;;---------------------------------------------------------------------|

(DEFUN BlockGet	(/ PiskSS 1StEntSS)
  (PROMPT "Select Title block to update:")
  (SETQ	PickSS	 (SSGET)		;get a selection set form the screen
	1StEntSS (ENTGET (SSNAME PickSS 0))
					;gets the entirety from the first
					;object in the sslist
	SubEnt	 (ASSOC -1 1StEntSS)	;gets the name for the object
  )					;Close SETQ
  SubEnt				;passes the value back
)					;Close DEFUN


;;;---------------------------------------------------------------------|
;;;	Sub-routine: CRSplit 						|
;;;---------------------------------------------------------------------|
;;;  Description: Tests for valid stings to split.             		|
;;;---------------------------------------------------------------------|
;;;  Sub-routine to look at the text stings that have been passwd to it |
;;;  to see if they are valid stings to split. The oly two stings that  |
;;;  are vlaid are the PROJECT_DECRIPTION and STREET_ADDRESS\ZIP	|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;			 	None					|
;;;---------------------------------------------------------------------|

(DEFUN CRSplit (/)
  (IF (NOT (OR (= (CDR DefChk) "PROJECT_DESCRIPTION")
	       (= (CDR DefChk) "ADDRESS_STREET")
	       (= (CDR DefChk) "CITY")
	   )
      )					;IF DefChk dose not equal "PROJECT_DESCREPTION" or "ADDRESS_STREET"...
    (ALERT
      "This is not a supported string to split.\nOnly the Description and the Address string can have CR in them."
    )					;THEN alert the user with this message
    (PROGN
      (IF (= (CDR DefChk) "PROJECT_DESCRIPTION")
					;IF the data of DefChk is the "PROJECT_DESCREPTION" ...
	(ProjDesSplit)			;THEN call the ProjDefSplit
	(AddSplit)			;ELSE call AddSplit
      )					;Close IF
    )					;Close PROGN
  )					;Close IF
)					;Close DEFUN


;;;---------------------------------------------------------------------|
;;;	Sub-routine: ProjDesSplite					|
;;;---------------------------------------------------------------------|
;;;  Description: Used to split the PROJECT_DESCRIPTION sting 		|
;;;---------------------------------------------------------------------|
;;;  Sub-routine to split the PROJECT_DESCRIPTION sting into a number of|
;;;  strings. If there are more than 4 strings then the user is warned	|
;;;  via the command line that only the first four are shown.		|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;   ProjDesStr	    ;; this is the sting that is taken for the  |
;;;			    :: data of AttDataNew			|
;;;   CDIndex		    ;; the location of the "\r\n" string within |
;;;			    ;; the ProjDesStr string			|
;;;   CRPlace		    ;; a list of were the "\r\n" stings are     |
;;;			    ;; in the ProjDesStr string			|
;;;   StringPD		    ;; section of the ProjDesStr string		|
;;;   StringLS		    ;; list of derived StringLS strings		|
;;;   PDloopCount	    ;; controler for number of times to loop	|
;;;---------------------------------------------------------------------|

(DEFUN ProjDesSplit(/  ProjDesStr   CRIndex
		     CRPlace	  StringPD     StringLS
		     PDLoopCount
		    )
  (SETQ	ProjDesStr (CDR AttDataNew)	;set ProjDesStr to the data of AttDataNew
	CRIndex	   (VL-STRING-SEARCH "\r\n" ProjDesStr)
					;set CRIndex to the position of the "\r\n" from the string ProjDesStr
	CRPlace	   (LIST CRIndex)	;set CRPlece to CRIndex in a list form
	StringPD   (SUBSTR ProjDesStr 1 CRIndex)
					;set StringPD to the valuse of ProjDesStr from the first chriter to were the "\r\n" is
	StringLS   (LIST StringPD)	;set StringLS to the data of StringPD and places it in a list format
  )					;Close SETQ
  (WHILE (NOT (= CRIndex nil))		;set a WHILE loop as long as CRIndex is not nil
    (SETQ ProjDesStr (SUBSTR ProjDesStr (+ CRIndex 3))
					;set ProjDesStr to the remainder of the ProjDesStr from the first "\r\n"
	  CRIndex    (VL-STRING-SEARCH "\r\n" ProjDesStr)
					;set CRIndex to the position of "\r\n" in the ProjDesStr
	  StringPD   (SUBSTR ProjDesStr 1 CRIndex)
					;set StringPD to the valuse of ProjDesStr from the first chriter to were the "\r\n" is
	  StringLS   (APPEND StringLS (LIST StringPD))
					;adds the value of StingPD to the end of StingLS
    )					;Close SETQ
    (IF	(NOT (= CRIndex nil))		;IF the value of CRIndex is not nil ...
      (SETQ CRPlace (APPEND CRPlace (LIST CRIndex)))
					;THEN Append the the value of CRIndex, in list form to the CRPlace list
    )					;Close IF
  )					;Close WHILE
  (IF (> (LENGTH StringLS) 4)		;IF the length of the StingLS is greater than 4 ...
    (PROGN				;Start THEN PROGN
      (PRINC "\There are ")
      (PRINC (LENGTH StringLS))
      (PRINC
	" lines to the Job Details Description field.
	\nOnly the first four (4) will be shown in the title block."
      )					;print on the command line that the StringLS list is greater then 4 and that only the first four will be shown
      (SETQ StringLS (REVERSE StringLS));set StringLS to its reverse order
      (WHILE (> (LENGTH StringLS) 4)	;WHILE the lenght of StringLS is greater than 4 ...
	(SETQ StringLS (CDR StringLS))	;set StringLS to all but the first item in the StingLS list
      )					;Close WHILE
      (SETQ StringLS (REVERSE StringLS));set StringLS to its reverse order
    )					;Close PROGN
  )					;Close IF
  (SETQ StringLS (REVERSE StringLS))
  (SETQ PDLoopCount (1- (LENGTH StringLS)))
					;set PDLoopCount to 1 less the length of StingLS
  (WHILE (>= PDLoopCount 0)		;WHILE PDLoopCount is greater then or equal to 0
    (SETQ AttDataNew
	   (CONS (CAR AttDataEx)
		 (NTH PDLoopCount StringLS)
					;Sets the AttDataNew to the constructed value of AttDataEx, DefChk and DWGPropData
	   )				;Close CONS
    )					;Close SETQ
    (ModAtt)				;call the sub-routine ModAtt
    (SETQ AttName (ENTNEXT AttName))	;set AttName to the next objectname
    (SETQ AttDataEx (ASSOC 1 (ENTGET AttName)))
					;set AttDataEx to the the 1 item of the data list of AttName
    (SETQ PDLoopCount (1- PDLoopCount))	;set PDLoopCount to 1 less
  )					;Close WHILE
)					;Close DEFUN


;;;---------------------------------------------------------------------|
;;;	Function: AddSplit						|
;;;---------------------------------------------------------------------|
;;;  Description: Used to warn the user the sting is not supported	|
;;;---------------------------------------------------------------------|
;;;  Sub-routine that warns the user that the they have stings, other   |
;;;  than the PROJECT_DESCRIPTION AND THE STREET_ADDRESS\ZIP that have	|
;;;  new line escape codes in them.					|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;   AddCR		    ;; string from AttDataNew			|
;;;   CDIndex		    ;; the location of the "\r\n" string within |
;;;			    ;; the ProjDesStr string			|
;;;   UsrZip		    ;; the postcode for the data that will be   |
;;;   			    ;; commbined with the street address	|
;;;   OldZip		    ;; existing data from the DWGPropData	|
;;;   			    ;; that will be replaced by UsrZip		|
;;;---------------------------------------------------------------------|

(DEFUN AddSplit	(/ AddCR CRIndex UsrZip OldZip)
  (SETQ AddCR (CDR AttDataNew))		;set AddCR to the data of AttDate
  (SETQ CRIndex (VL-STRING-SEARCH "\r\n" AddCR))
					;set CRIndex to the location of "\r\n" in the AddCR string
  (SETQ AttDataNew (SUBSTR AddCR 1 CRIndex))
					;set AttDataNew to the section of AddCR string from the first charactor to the location of the "\r\n"
  (SETQ AddCR (SUBSTR AddCR (+ CRIndex 3)))
					;set AddCR to the remainder of AddCR from the location of CRIndex + 3
  (SETQ AttDataNew (CONS 1 AttDataNew))	;set AttDataNew to a constructed list of 1 . AttDataNew
  (ModAtt)				;call sub-routine ModAtt
  (SETQ UsrZip (ASSOC 307 DWGPropData))	;set UsrZip to the item 307 of DWGPropData
  (SETQ UsrZip (CDR UsrZip))		;set UsrZip to the data section of UsrZip
  (SETQ UsrZip (VL-STRING->LIST UsrZip));set UsrZip to an ANSII charactor list
  (SETQ AddCR (VL-STRING->LIST AddCR))	;set AddCR to an ANSII charactor list
  (SETQ UsrZip (APPEND AddCR (LIST 32) UsrZip))
					;set UsrZip to a list that is appended of AddCR charactor 32 and UsrZip
  (SETQ UsrZip (VL-LIST->STRING UsrZip)) ;set UsrZip back to a string

;;;  Following lines commented out as they were only used for testing
;;;  (setq AttDataNew (CONS 1 UsrZip))
;;;  (setq AttName (entnext AttName))
;;;  (ModAtt)

  (SETQ UsrZip (CONS 307 UsrZip))	;set UsrZip to a constructed list of 307 . UsrZip
  (SETQ OldZip (ASSOC 307 DWGPropData))	;set OldZip to item 307 . DWGPropData
  (SETQ DWGPropData (SUBST UsrZip OldZip DWGPropData))
					;set DWGPropData to the list that has substatue UsrZip inplace of OldZip in the list DWGPropData
)					;Close DEFUN


;;;---------------------------------------------------------------------|
;;;	Function: WarnSplit						|
;;;---------------------------------------------------------------------|
;;;  Description: Used to warn the user the sting is not supported	|
;;;---------------------------------------------------------------------|
;;;  Sub-routine that warns the user that the they have stings, other   |
;;;  than the PROJECT_DESCRIPTION AND THE STREET_ADDRESS\ZIP that have	|
;;;  new line escape codes in them.					|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;   WarnStr		   ;; sting telling the users that there is     |
;;;			   ;; more data than can be displayed		|
;;;---------------------------------------------------------------------|

(DEFUN WarnSplit (/ WarnStr)
  (SETQ	WarnStr
	 (STRCAT			;Creates a string from DefChk
	   "The \""
	   (CDR DefChk)			;Gets the Attribute name
	   "\" string has escape codes in it.\nThe string is as follows:\n"
	   (CDR (ASSOC (CAR DefChk) DwgPropData))
					;Strips the sting with the "\r\n" string in it
	 )				;Close STRCAT
  )					;Close SETQ
  (ALERT WarnStr)			;Shows a warning box with the information about the string
  (ModAtt)
)					;Close DEFUN


;;;---------------------------------------------------------------------|
;;;	Function: ModAtt 						|
;;;---------------------------------------------------------------------|
;;;  Description: Used to update blcok attribute data         		|
;;;---------------------------------------------------------------------|
;;;  Sub-routine to update the attribute date in a block. Takes		|
;;;  information from passed to it and applyeds it. After doing this	|
;;;  ENTMOD is used to apply it back to the drawing database.		|
;;;---------------------------------------------------------------------|
;;;  The variables are as follows:					|
;;;   xxxxxx		    ;; 						|
;;;---------------------------------------------------------------------|

(DEFUN ModAtt (/)
  (SETQ	AttDataNew
	 (SUBST AttDataNew AttDataEx (ENTGET AttName))
  )					;Sets AttDataNew to the substituted values of AttDataNew into AttName
  (ENTMOD AttDataNew)			;Applies the new values to the drawing data base.
)