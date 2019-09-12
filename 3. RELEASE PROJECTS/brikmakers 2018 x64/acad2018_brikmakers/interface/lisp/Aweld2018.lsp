;------------------------------------------------------------------------
;                    AWELD  Ver. 4.0
;          
;          Welding Symbol Placement Procedure
;
;         By Eric Howard [76424,1036] 11-18-92
;
;------------------------------------------------------------------------
;
;          The AWELD Symbol Placement Procedure draws welding symbols
;     without using a library of blocks. All weld symbol geometry is placed
;     on the layer "WELD_SYM" with the blue color assigned to the layer.
;     The size is determined by the variable "DIMSCALE".
;
;     The ACAD.MNU file must be modified for maximum
;     productive. To do this the following lines of code should
;     be added.
;
;     Add this line to any popup menu:
;
;     [Weld Symbol]^C^C(IF (NULL C:AWELD) (PROGN (PRINC "Initial load -- Please wait...")
;                      (PRINC) (LOAD "{path name} AWELD")(PRINC))(PRINC)) AWELD
;
;     The {path name} is where the LISP file is located
;     ie... "C:/AUTODESK/LISP". Use forward slash instead
;     of the back slash. This line allows that if the
;     program is already loaded that it will not load again.
;
;     To manually load the file "AWELD.LSP" type the following
;     at the command prompt: (load "{path name}/AWELD") and
;     type AWELD at the command prompt to start the program.
;
;     The file "AWELD.DCL" is the dialog box that selects the
;     weld type and must be placed in a directory where AUTOCAD
;     can find it. The file "awelds.SLB" is the slide library 
;     for the dialog box and must also be placed in a directory 
;     where AUTOCAD can access it.
;     
;     This version of AWELD.LSP has been written for Release 12. It
;     will not work in any previous version. 
;
;
(defun c:aweld (/ OLDERROR ORTH LYR LR1 SCL P1 P2 A1 P3 TP1 TP2 TXTSTY)
    ;
    ;---------------------------------
    ;   Set Environment Information
    ;---------------------------------
    ;
    (setvar "cmdecho" 0)
    (setq olderror *error* *error* welderror)
    (setq ORTH (getvar "orthomode"))
    (setq TXTSTY (getvar "textstyle"))
    (setvar "orthomode" 0)
    (setq SCL (getvar "dimscale"))
    (setq LYR (getvar "clayer"))
    (setq MRR (getvar "mirrtext"))
    (setq MIRSET (ssadd))
    (setvar "mirrtext" 0)
    ;
    ;----------------------------------------------------
    ;   Setting Layer Preliminary Information
    ;----------------------------------------------------
    ;
       
	(if (= (tblsearch "layer" "dim") nil)
		(command "layer" "new" "dim" "")
		(command "layer" "set" "dim" "")
	)
    ;
    ;------------------------------------------
    ;   Select Points and Symbol Information
    ;------------------------------------------
    ;
    (prompt "\nSelect weld joint : ")
;;;    (while (setq P1 (getpoint))
;;;           (setq P2 (getpoint P1 "\nNext point"))

      		(setq P1 (getpoint))
           (setq P2 (getpoint P1 "\nNext point : "))
        (wset)
      
;;; (setq WELDTYP "(plweld)")
;;; (setq wt_mode "pl_weld")

;;;	(action_tile "sq_weld" "(setq WELDTYP \"(sqweld)\")")
;;;	(action_tile "u_weld" "(setq WELDTYP \"(uweld)\")")
;;;	(action_tile "j_weld" "(setq WELDTYP \"(jweld)\")")
;;;	(action_tile "v_weld" "(setq WELDTYP \"(vweld)\")")
;;;	(action_tile "b_weld" "(setq WELDTYP \"(bweld)\")")
;;;	(action_tile "flv_weld" "(setq WELDTYP \"(flvweld)\")")
;;;	(action_tile "flb_weld" "(setq WELDTYP \"(flbweld)\")")
;;;	(action_tile "fil_weld" "(setq WELDTYP \"(filweld)\")")
;;;	(action_tile "sur_weld" "(setq WELDTYP \"(surweld)\")")
;;;	(action_tile "fle_weld" "(setq WELDTYP \"(fleweld)\")")
;;;	(action_tile "flc_weld" "(setq WELDTYP \"(flcweld)\")")
;;;	(action_tile "pl_weld" "(setq WELDTYP \"(plweld)\")")
      
        (if (= WELDTYP nil)
            (exit)
        )    
        (if (or (<= (angle P1 P2) (dtr 90)) (>= (angle P1 P2) (dtr 270)))
            (setq A1 0)
            (setq A1 (dtr 180))
        )
        (setq P3 (polar P2 A1 (* SCL 12.00)))
        ;
        ;-----------------------------------------------------------
        ;   Sets text placement points
        ;-----------------------------------------------------------
        ;
        (if (= 0 A1)
            (progn
                (setq TP1 (polar P2 A1 (* SCL 6.00)))
                (setq TP2 (polar P2 A1 (* SCL 21.00)))
            )
            (progn
                (setq TP1 (polar P2 A1 (* SCL 21.00)))
                (setq TP2 (polar P2 A1 (* SCL 6.00)))
            )
        )    
        (setq TP1 (polar TP1 (dtr 90) (* SCL 3.00)))        
        (setq TP2 (polar TP2 (dtr 90) (* SCL 3.00)))     
        (command "pline" P1 "w" 0 (* SCL 1.25) (polar P1 (angle P1 P2) (* SCL 3.0)) "w" 0 0 P2 (polar P2 A1 (* SCL 20.0)) "")
        ;
        ;-------------------------------------------------
        ;   "All Around symbol" information if required
        ;-------------------------------------------------
        ;
        (if (= ALLR "1")    
            (command "circle" P2 (* SCL 2.00))
        )
        ;
        ;-------------------------------------------------
        ;   Invokes program to draw selected weld type
        ;-------------------------------------------------
        ;
	    (setq LR1 (tblsearch "layer" "3t"))
	    (if (= LR1 nil)
	        (command "layer" "new" "3t" "color" "YELLOW" "3t" "set" "3t" "")
	        (command "layer" "set" "3t" "")
	    )    
        (if (= WELDTYP2 "(baweld)")
            (eval (read WELDTYP2))
        )
        (eval (read WELDTYP))
        (if (= (cdr (assoc 40 (tblsearch "style" TXTSTY))) 0)
            (progn
                (command "text" "j" "mc" TP1 (* SCL 3.5) 0 WELDSIZE)
                (mir_set)
                (command "text" "j" "mc" TP2 (* SCL 3.5) 0 PITCH)
                (mir_set)
            )
            (progn
                (command "text" "j" "mc" TP1 0 WELDSIZE)
                (mir_set)
                (command "text" "j" "mc" TP2 0 PITCH)
                (mir_set)
            )
        )    
        (if (= WELDSIDE "This")
            (setq QUES "Y")
            (setq QUES "N")
        )
        (if (or (= WELDSIDE "This") (= WELDSIDE "Both"))
            (cond
                ((/= WELDTYP2 "(baweld)")(command "mirror" MIRSET "" P2 (polar P2 A1 0.5) QUES))
                ((= WELDTYP2 "(baweld)")(command "mirror" MIRSET "r" RBAK "" P2 (polar P2 A1 0.5) QUES))
            )
        )

        ;
        ;--------------------------------------------
        ;   "Field symbol" information if required
        ;--------------------------------------------
        ;
        (if (= FIELDW "1")
            (flag)
        )    
        :
        ;
        (notes2)
        ;
        ;--------------------------------------------
        ;   Loop to restart weld symbol placement
        ;--------------------------------------------
        ;



  
;;;        (command "layer" "set" "dim" "")
;;;        (setq MIRSET (ssadd))
;;;        (prompt "\nSelect weld joint <Return to exit>")
  
;;;    )

  

    ;
    ;-------------------
    ;   Exit Procedure
    ;-------------------
    ;
    (setvar "orthomode" ORTH)
    (setvar "mirrtext" MRR)
    (command "layer" "set" LYR "")
    (setq *error* OLDERROR)
    (princ)    
)
;
;--------------------------------------------
;   Procedure to set defaults
;--------------------------------------------
;
;
;--------------------------------------------
;   Error handler
;--------------------------------------------
;
(defun welderror (msg)
    (setvar "orthomode" ORTH)
    (setvar "mirrtext" MRR)
    (command "layer" "set" LYR "")
    (setq *error* OLDERROR)
    (setq msg1 msg)
    (if (= msg "quit / exit abort")
        (prompt "\nNo Weld Type was Selected\n")
    )
    (princ)
)
;
;************************************
;   Fillet Weld Symbol
;************************************
;
(defun filweld ()
    (command "pline" (polar P3 (dtr 180) (* SCL 2.50)) (polar (getvar "lastpoint") (dtr 90) (* SCL 5.00)) (polar P3 0 (* SCL 2.50)) "")
    (mir_set)
)
;
;***********************************
;   Square Weld Symbol
;***********************************

(defun sqweld ()
    (command "pline" (polar P3 0 (* SCL 1.5)) (polar (getvar "lastpoint") (dtr 90) (* SCL 6.0)) "")
    (mir_set)
    (command "pline" (polar P3 (dtr 180) (* SCL 1.5)) (polar (getvar "lastpoint") (dtr 90) (* SCL 6.0)) "")
    (mir_set)
)
;
;*********************************
;   "U" Weld Symbol
;*********************************
;
(defun uweld (/ UPT)
    (setq UPT (polar P3 (dtr 90) (* SCL 3.0)))
    (command "line" P3 UPT "")
    (mir_set)
    (command "arc" (polar UPT (dtr 45) (* SCL 4.5)) UPT (polar UPT (dtr 135) (* SCL 4.5)))
    (mir_set)
)
;
;*********************************
;   "J" Weld Symbol OTHER Side
;*********************************
;
(defun jweld (/ JPT) 
    (setq JPT (polar P3 (dtr 90) (* SCL 6.0)))
    (command "line" P3 JPT "")
    (mir_set)
    (command "arc" (polar P3 (dtr 90) (* SCL 3.0)) "c" JPT (polar JPT 0 (* SCL 3.0)))
    (mir_set)
)
;
;*********************************
;   "V" Weld Symbol
;*********************************
;
(defun vweld ()
    (command "line" P3 (polar P3 (dtr 60) (* SCL 5.0)) "")
    (mir_set)
    (command "line" P3 (polar P3 (dtr 120) (* SCL 5.0)) "")
    (mir_set)
)
;
;*********************************
;   Bevel Weld Symbol
;*********************************
;
(defun bweld ()
    (command "line" P3 (polar P3 (dtr 45) (* SCL 7.0)) "")
    (mir_set)
    (command "line" P3 (polar P3 (dtr 90) (* SCL 5.0)) "")
    (mir_set)
)
;
;************************************
;   Flare-V Weld Symbol
;************************************
;
(defun flvweld ()
    (command "arc" "c" (polar P3 0 (* SCL 4.5)) (polar (getvar "lastpoint") (dtr 90) (* SCL 3.0)) (polar P3 0 (* SCL 1.5)))
    (mir_set)
    (command "arc" (polar P3 (dtr 180) (* SCL 1.5)) "c" (polar P3 (dtr 180) (* SCL 4.5)) (polar (getvar "lastpoint") (dtr 90) (* SCL 3.0)))
    (mir_set)
)
;
;****************************************
;   Flare Bevel Weld Symbol
;****************************************
;
(defun flbweld ()
    (command "arc" "c" (polar P3 0 (* SCL 4.5)) (polar (getvar "lastpoint") (dtr 90) (* SCL 3.0)) (polar P3 0 (* SCL 1.5)))
    (mir_set)
    (command "pline" (polar p3 (dtr 180) (* SCL 1.5)) (polar (getvar "lastpoint") (dtr 90) (* SCL 3.0)) "")
    (mir_set)
)
;

;*************************
;   Surface Weld Symbol
;*************************
;
(defun surweld ()
    (command "arc" (polar P3 0 (* SCL 3.0)) "C" (polar P3 0 (* SCL 1.5)) P3)
    (mir_set)
    (command "arc" P3 "c" (polar P3 (dtr 180) (* SCL 1.5)) (polar P3 (dtr 180) (* SCL 3.0)))
    (mir_set)
)
;
;***************************************
;   Flange Edge Weld Symbol
;***************************************
;
(defun fleweld ()
    (command "arc" (polar P3 (dtr 180) (* SCL 4.5)) "c" (polar (getvar "lastpoint") (dtr 90) (* SCL 3.0)) (polar (getvar "lastpoint") 0 (* SCL 0.12)))
    (mir_set)
    (command "pline" (getvar "lastpoint") (polar (getvar "lastpoint") (dtr 90) (* SCL 3.0)) "")
    (mir_set)
    (command "pline" (polar (getvar "lastpoint") 0 (* SCL 3.0)) (polar (getvar "lastpoint") (dtr 270) (* SCL 3.0)) "")
    (mir_set)
    (command "arc" (getvar "lastpoint") "c" (polar (getvar "lastpoint") 0 (* SCL 3.0)) (polar P3 0 (* SCL 4.5)))
    (mir_set)
)
;
;***************************************
;   Flange Corner Weld Symbol
;***************************************
;
(defun flcweld ()
    (command "pline" (polar P3 (dtr 180) (* SCL 1.5)) (polar (getvar "lastpoint") (dtr 90) (* SCL 6.0)) "")
    (mir_set)
    (command "pline" (polar (getvar "lastpoint") 0 (* SCL 3.0)) (polar (getvar "lastpoint") (dtr 270) (* SCL 3.0)) "")
    (mir_set)
    (command "arc" (getvar "lastpoint") "c" (polar (getvar "lastpoint") 0 (* SCL 3.0)) (polar P3 0 (* SCL 4.5)))
    (mir_set)
)
;
;********************************
;   Plug Weld Symbol
;********************************
;
(defun plweld ()
    (command "pline" (polar P3 (dtr 180) (* SCL 6.0)) (polar (getvar "lastpoint") (dtr 90) (* SCL 6.0)) (polar (getvar "lastpoint") 0 (* SCL 6.0)) (polar P3 0 (* SCL 0)) "")
    (mir_set)
)
;
;*************************************
;   Backing  Weld Symbol
;*************************************
;
(defun baweld ()
    (if (= WELDSIDE "This")
        (command "arc" "c" P3 (polar P3 0 (* SCL 1.5)) (polar p3 (dtr 180) (* SCL 1.5)))
        (command "arc" "c" P3 (polar P3 (dtr 180) (* SCL 1.5)) (polar P3 0 (* SCL 1.5)))
    )
    (setq RBAK (entlast))
)
;
;********************************************
;   "Field symbol" information if required
;********************************************
;
(defun flag (/ FPT)
    (setq fpt (polar P2 (dtr 90) (* SCL 8.0)))
    (setq fpt (polar FPT A1 (* SCL 5.0)))
    (command "line" P2 (polar P2 (dtr 90) (* SCL 9.5)) "")
    (command "solid" (getvar "lastpoint") FPT (polar P2 (dtr 90) (* SCL 6.5)) "")
    (command)
)
;
;********************************************
;   Converts degrees to radians
;********************************************
;
(defun dtr (ANG)(* PI (/ ANG 180.0)))
;
;
;
;
;
;*************************************
;   Places Notes and Note Tail
;*************************************
;
;
(defun notes2 ()    
    (setq TXTPT (polar P2 A1 (* SCL 23.0)))
    (setq TAILPT (polar P2 A1 (* SCL 20.0)))
    (setq PASS 1)
    (setq RPT 0)
    (repeat 3
        (if (/= (eval (read (strcat "note" (itoa PASS)))) "")
            (setq RPT (1+ RPT))
        )
        (setq PASS (1+ PASS))
    )        
    (if (/= RPT 0)                                             
        (progn
            (cond
                ((= RPT 1)(setq TXTPT1 TXTPT)(cond ((/= NOTE1 "")(setq NTE1 NOTE1))
                        ((/= NOTE2 "")(setq NTE1 NOTE2))
                        ((/= NOTE3 "")(setq NTE1 NOTE3))
                    )
                )                                   
                ((= RPT 2)(setq TXTPT (polar TXTPT A1 (* SCL 1.6)))(setq TXTPT1 (polar TXTPT (dtr 90) (* SCL 2.0)))(cond ((= NOTE1 "")(setq NTE1 NOTE2 NTE2 NOTE3))
                        ((= NOTE2 "")(setq NTE1 NOTE1 NTE2 NOTE3))
                        ((= NOTE3 "")(setq NTE1 NOTE1 NTE2 NOTE2))
                    )
                )                                                                                                               
                ((= RPT 3)(setq TXTPT (polar TXTPT A1 (* SCL 3.0)))(setq TXTPT1 (polar TXTPT (dtr 90) (* SCL 4.0)))(setq NTE1 NOTE1 NTE2 NOTE2 NTE3 NOTE3))         
            )
            (setq PASS 1)
            (if (= A1 0)
                (progn
                    (repeat RPT
                        (if (= (cdr (assoc 40 (tblsearch "style" TXTSTY))) 0)
                            (command "text" "j" "ml" TXTPT1 (* SCL 3.5) 0 (eval (read (strcat "NTE" (itoa PASS)))))
                            (command "text" "j" "ml" TXTPT1 0 (eval (read (strcat "NTE" (itoa PASS)))))
                        )
                        (setq PASS (1+ PASS))
                        (setq TXTPT1 (polar TXTPT1 (dtr 270) (* SCL 4.0)))
                    )
                )
                (progn
                    (repeat RPT
                        (if (= (cdr (assoc 40 (tblsearch "style" TXTSTY))) 0)
                            (command "text" "j" "mr" TXTPT1 (* SCL 2.5) 0 (eval (read (strcat "NTE" (itoa PASS)))))
                            (command "text" "j" "mr" TXTPT1 0 (eval (read (strcat "NTE" (itoa PASS)))))
                        )
                        (setq PASS (1+ PASS))
                        (setq TXTPT1 (polar TXTPT1 (dtr 270) (* SCL 4.0)))
                    )    
                )    
            )
            (if (= A1 0)
                (command ".line" TAILPT (polar TAILPT (dtr 45) (* SCL 7.0)) "" "line" TAILPT (polar TAILPT (dtr -45) (* SCL 7.0)) "")
                (command ".line" TAILPT (polar TAILPT (dtr 135) (* SCL 7.0)) "" "line" TAILPT (polar TAILPT (dtr -135) (* SCL 7.0)) "")
            )
        )
    )
)
;
;
;*********************************************
;   Sets up and Calls Setup Dialog Box
;*********************************************
;
;
;
(defun wset ()
    (setq dcl_id (load_dialog "C:\\ACAD2018_BRIKMAKERS\\Interface\\Lisp\\aweld.DCL"))
    (if (not (new_dialog "aweld" dcl_id))
        (exit)
    )
    (setq x (dimx_tile "fil_weld")
        y (dimy_tile "fil_weld")
    )      
    (start_image "fil_weld")
    (slide_image 0 0 x y "awelds(filweld)")
    (end_image)
    (setq x (dimx_tile "sq_weld")
        y (dimy_tile "sq_weld")
    )      
    (start_image "sq_weld")
    (slide_image 0 0 x y "awelds(sqweld)")
    (end_image)
    (setq x (dimx_tile "u_weld")
        y (dimy_tile "u_weld")
    )      
    (start_image "u_weld")
    (slide_image 0 0 x y "awelds(uweld)")
    (end_image)                                            
    (setq x (dimx_tile "j_weld")                           
        y (dimy_tile "j_weld")                             
    )                                                      
    (start_image "j_weld")                                 
    (slide_image 0 0 x y "awelds(jweld)")       
    (end_image)                                            
    (setq x (dimx_tile "v_weld")                           
        y (dimy_tile "v_weld")                             
    )                                                      
    (start_image "v_weld")                                 
    (slide_image 0 0 x y "awelds(vweld)")        
    (end_image)
    (setq x (dimx_tile "b_weld")
        y (dimy_tile "b_weld")
    )      
    (start_image "b_weld")
    (slide_image 0 0 x y "awelds(bweld)")
    (end_image)
    (setq x (dimx_tile "flv_weld")
        y (dimy_tile "flv_weld")
    )      
    (start_image "flv_weld")
    (slide_image 0 0 x y "awelds(flvweld)")
    (end_image)
    (setq x (dimx_tile "flb_weld")
        y (dimy_tile "flb_weld")
    )      
    (start_image "flb_weld")
    (slide_image 0 0 x y "awelds(flbweld)")
    (end_image)

    (setq x (dimx_tile "sur_weld")
        y (dimy_tile "sur_weld")
    )      
    (start_image "sur_weld")                                
    (slide_image 0 0 x y "awelds(surweld)")    
    (end_image)                                             
    (setq x (dimx_tile "fle_weld")                          
        y (dimy_tile "fle_weld")                            
    )                                                       
    (start_image "fle_weld")                                
    (slide_image 0 0 x y "awelds(fleweld)")    
    (end_image)                                             
    (setq x (dimx_tile "flc_weld")                          
        y (dimy_tile "flc_weld")                            
    )                                                       
    (start_image "flc_weld")
    (slide_image 0 0 x y "awelds(flcweld)")
    (end_image)
    (setq x (dimx_tile "pl_weld")
        y (dimy_tile "pl_weld")
    )      
    (start_image "pl_weld")
    (slide_image 0 0 x y "awelds(plweld)")
    (end_image)
    (aw_deflt)
    (action_tile "ba_weld" "(setq BA_DEF $value)(ba_set)")
    (action_tile "ar_weld" "(setq ALLR $value)")
    (action_tile "fi_weld" "(setq FIELDW $value)")
    (action_tile "sq_weld" "(setq WELDTYP \"(sqweld)\")")
    (action_tile "u_weld" "(setq WELDTYP \"(uweld)\")")
    (action_tile "j_weld" "(setq WELDTYP \"(jweld)\")")
    (action_tile "v_weld" "(setq WELDTYP \"(vweld)\")")
    (action_tile "b_weld" "(setq WELDTYP \"(bweld)\")")
    (action_tile "flv_weld" "(setq WELDTYP \"(flvweld)\")")
    (action_tile "flb_weld" "(setq WELDTYP \"(flbweld)\")")
    (action_tile "fil_weld" "(setq WELDTYP \"(filweld)\")")
    (action_tile "sur_weld" "(setq WELDTYP \"(surweld)\")")
    (action_tile "fle_weld" "(setq WELDTYP \"(fleweld)\")")
    (action_tile "flc_weld" "(setq WELDTYP \"(flcweld)\")")
    (action_tile "pl_weld" "(setq WELDTYP \"(plweld)\")")
    (action_tile "this_side" "(setq WELDSIDE \"This\")")
    (action_tile "other_side" "(setq WELDSIDE \"Other\")")
    (action_tile "both_side" "(setq WELDSIDE \"Both\")")
    (action_tile "fill_size" "(setq WELDSIZE $value)")
    (action_tile "pitch_size" "(setq PITCH $value)")
    (action_tile "note1" "(setq NOTE1 $value)")
    (action_tile "note2" "(setq NOTE2 $value)")
    (action_tile "note3" "(setq NOTE3 $value)")
    (action_tile "accept" "(done_dialog)")
    (action_tile "cancel" "(exit)")
    (start_dialog)
    (unload_dialog dcl_id)
)    
;
;
;***********************************************
;   Determines if Backing Weld is to be Used
;***********************************************
;
;
(defun ba_set ()
    (mode_tile "both_side" (atoi $value))
    (if (and (= "1" $value)(= (get_tile "both_side") "1"))
        (progn
            (set_tile "this_side" "1")
            (setq WELDSIDE "This")
        )
    )
    (if (= "1" $value) 
        (setq WELDTYP2 "(baweld)")
        (setq WELDTYP2 nil)
    )
)
;
;
;***********************************************
;   Sets the Defaults in the Setup Dialog Box
;***********************************************
;
;
(defun aw_deflt ()
    (if (cond ((= WELDTYP "(sqweld)")(setq wt_mode "sq_weld"))
            ((= WELDTYP "(uweld)")(setq WT_MODE "u_weld"))
            ((= WELDTYP "(jweld)")(setq WT_MODE "j_weld"))
            ((= WELDTYP "(vweld)")(setq WT_MODE "v_weld"))
            ((= WELDTYP "(bweld)")(setq WT_MODE "b_weld"))
            ((= WELDTYP "(flvweld)")(setq WT_MODE "flv_weld"))
            ((= WELDTYP "(flbweld)")(setq WT_MODE "flb_weld"))
            ((= WELDTYP "(filweld)")(setq WT_MODE "fil_weld"))
            ((= WELDTYP "(surweld)")(setq WT_MODE "sur_weld"))
            ((= WELDTYP "(fleweld)")(setq WT_MODE "fle_weld"))
            ((= WELDTYP "(flcweld)")(setq WT_MODE "flc_weld"))
            ((= WELDTYP "(plweld)")(setq WT_MODE "pl_weld"))
        )
        (mode_tile WT_MODE 2)
    )
    (if (cond ((= WELDSIDE "This")(setq WS_MODE "this_side"))
            ((= WELDSIDE "Other")(setq WS_MODE "other_side"))
            ((= WELDSIDE "Both")(setq WS_MODE "both_side"))
        )
        (set_tile ws_mode "1")
    )    
    (if (= WELDSIZE nil)
        (setq WELDSIZE "")
    )
    (set_tile "fill_size" WELDSIZE)
    (if (= PITCH nil)
        (setq PITCH "")
    )
    (set_tile "pitch_size" PITCH)
    (if (= NOTE1 nil)
        (setq NOTE1 "")
    )
    (set_tile "note1" NOTE1)
    (if (= NOTE2 nil)
        (setq NOTE2 "")
    )
    (set_tile "note2" NOTE2)
    (if (= NOTE3 nil)
        (setq NOTE3 "")
    )
    (set_tile "note3" NOTE3)
    (if (= BA_DEF "1")
        (progn
            (set_tile "ba_weld" "1")
            (mode_tile "both_side" 1)
        )
    )
    (if (= ALLR "1")
        (set_tile "ar_weld" "1")
    )        
    (if (= fieldw "1")
        (set_tile "fi_weld" "1")
    )
)
;
;
;**********************************************
;   Creates the Selection Set to be Mirrored
;**********************************************
;
;
(defun mir_set ()
    (ssadd (entlast) MIRSET)
)
