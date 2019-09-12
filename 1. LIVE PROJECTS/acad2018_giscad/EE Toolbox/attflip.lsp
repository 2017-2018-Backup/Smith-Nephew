;This function checks the rotation angle of attributes of a given block and if 
;the rot angle is not readable to the screen, it flips and mirrors the attributes
;to be always readable.  It has the smarts to flip logical groups of attributes
;as well depending on passed parameters.  Program can be used in Surveyors view 
;(view rotated 90 deg off zero) or regular view.

;Written by David Noble - July 1998

;Copyright 2000 MW Technologies  All Rights Reserved
;www.mwtech.com

;''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
; You are free to use this code within your own applications,
; but you are expressly forbidden from selling or otherwise 
; distributing this source code without prior written consent.
; This includes both posting free demo projects made from this 
; code as well as reproducing the code in text or html format. 
;''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

;MW Technologies provides this program "As Is" and with all faults.  MW Technologies
;specifically disclaims any implied warranty of fitness for a particular use.
;MW Technologies does not warrant that the operation of the program will be
;uninterrupted or error free.


(defun _AttFlip (e1 setnum attnum / attcnt attlst cnt cnt1 cnt3 dif dis e2 elst etyp inspt newlst  
                                     newpt newpt_x newpt_y rot str txtrot txtrot1 tmplst vflag)
  
  ;e1     = ename of block with attributes to check
  ;setnum = number of sets of attributes : Meaning, if a block has 12 attributes and the attributes are
  ;                                        grouped logically in sets of 3 (ex. BH#, Dip and length) then
  ;                                        there would be 4 sets.  This allows logical group flipping. <Int>
  ;                                        When there is logical grouping, the attribute tags must have a number
  ;                                        at the end of the string designating which group they are in.
  ;                                        (Ex. BH1, Dip1, Len1, BH2, Dip2, Len2 etc...)
  ;attnum = number of attributes in each setnum set. : Int (In the example above it would be 3 )
  
  
  ;NOTES: 1.  If a block has 6 individual attributes : setnum = 6 & attnum = 1  (no grouping)
  ;           (each tag must also end in a number designating position - 1, 2, 3, ...)
  ;       2.  The textbox function uses the txt style size if it's set to a non 0 #, or 
  ;           the variable "textsize" if the txt style size is set to zero.
  ;       3.  vflag variable = flag for surveyors view setup
  
  
  ;first get the insertion points of attributes (one set at a time - 1st <attnum>, then 2nd <attnum> etc.)

  (setq e2      e1
        elst    (entget e1)
        attcnt  0
        cnt1    1
        dif     1E-6
  )
  
  (if (equal (getvar "viewtwist") (/ (* 270.0 pi) 180.0)  dif)
    (setq vflag T)
  )
  
  (repeat setnum                 
   (while (and (< attcnt attnum) (/= etyp "SEQEND")) 
     (setq e1 (entnext e1))  ;bring up next entity in blk def
     (if e1
       (setq elst (entget e1)                
             etyp (cdr (assoc 0 elst)) 
       )
     )                                                  

     (if (= etyp "ATTRIB")  
      (progn
        (setq rot    (cdr (assoc 50 elst))
              txtsty (cdr (assoc 7 elst))
              txtsz  (cdr (assoc 40 elst))
        ) 
        (if vflag                                     ;if screen is set up using surveyors view - adjust rot
          (setq rot (abs (- rot (/ pi 2))))
        )

        (if (not (or (and (>= rot 0.0) (<= rot (/ pi 2))) (and (>= rot (* 1.5 pi)) (<= rot (* pi 2))))) 
          (if (> setnum 1)
            (if (= (substr  (cdr (assoc 2 elst)) (strlen (cdr (assoc 2 elst)))) (itoa cnt1))
              (setq inspt   (trans (cdr (assoc 10 elst)) e1 1)
                    str     (cdr (assoc 1 elst))
                    txtrot1 (cdr (assoc 50 elst))
                    attlst  (append attlst (list (list e1 inspt str)))
                    attcnt  (1+ attcnt)
              )
            )
          ;else
            (setq inspt   (trans (cdr (assoc 10 elst)) e1 1)
                  str     (cdr (assoc 1 elst))
                  txtrot1 (cdr (assoc 50 elst))
                  attlst  (append attlst (list (list e1 inspt str)))
                  attcnt  (1+ attcnt)
            )
          )

        )
       )
     )

   )

   (if attlst
     (if (equal (length attlst) attnum)
      (progn
       (setq cnt    0
             cnt3   (- (length attlst) 1)
       )
       
       (if (/= (/ (length attlst) 2.0) (fix (/ (length attlst) 2.0)))    ;if list length is ODD 
        (progn 
         (if (> (length attlst) 2) 
          (progn
           (repeat (/ (- (length attlst) 1) 2)     ;repeat to middle attribute position
             (setq tmplst (list (car (nth cnt attlst)) (cadr (nth cnt3 attlst)) (caddr (nth cnt attlst)))
                   newlst (append newlst (list tmplst))
                   cnt    (1+ cnt)
                   cnt3   (- cnt3 1)
             )
           )
         
           (setq tmplst (nth cnt attlst)              ;middle attribute
                 newlst (append newlst (list tmplst))
                 cnt    (1+ cnt)
                 cnt3   (- cnt3 1)
           )

           (repeat (/ (- (length attlst) 1) 2)     ;from middle attribute position to end
             (setq tmplst (list (car (nth cnt attlst)) (cadr (nth cnt3 attlst)) (caddr (nth cnt attlst)))
                   newlst (append newlst (list tmplst))
                   cnt    (1+ cnt)
                   cnt3   (- cnt3 1)
             )
           )
          )
         ;else its only 1 attribute
           (setq tmplst (car attlst) 
                 newlst (append newlst (list tmplst))
           )
         )
         
        )
       ;else its EVEN
         (repeat (length attlst) 
           (setq tmplst (list (car (nth cnt attlst)) (cadr (nth cnt3 attlst)) (caddr (nth cnt attlst)))
                 newlst (append newlst (list tmplst))
                 cnt    (1+ cnt)
                 cnt3   (- cnt3 1)
           )
         )
        )
       
       (setq attlst nil
             cnt    0
       )
       
       (setvar "textstyle" txtsty) ;Set default textsize and textstyle variables (used for textbox function)  
       (setvar "textsize" txtsz)   
       
       ;newlst = list of att properties after flipping 1st and last inspts, 2nd and 2nd to last inspts, etc, etc...

       (repeat (length newlst)
         (setq str     (list (cons 1 (caddr (nth cnt newlst))))
               inspt   (cadr (nth cnt newlst))
               newpt   (textbox str)                       ;find out dimensions of text
               newpt_x (car (cadr newpt)) 
               newpt_y (cadr (cadr newpt)) 
               txtrot  (+ txtrot1 (atan newpt_y newpt_x))  ;angle from txt inspt (bottom left [BL]) to txt upper right [UR]
               dis     (sqrt (+ (* newpt_x newpt_x) (* newpt_y newpt_y)))   ;distance from txt BL to txt UR
          )

          (setq newpt  (polar inspt txtrot dis)        ;txt UR point
                e1     (car (nth cnt newlst))
                attlst (append attlst (list (list e1 newpt)))
                cnt    (1+ cnt)
          )

        )
                                                       ;attlst = list of attribute enames with their new inspts.
        (setq cnt 0)
        (repeat (length attlst)
          (setq e1   (car (nth cnt attlst))
                elst (entget e1)
                elst (subst (cons 10 (cadr (nth cnt attlst))) (assoc 10 elst) elst)  ;new inspt
                elst (subst (cons 50 (+ txtrot1 pi)) (assoc 50 elst) elst)           ;new txt rotation angle
                cnt  (1+ cnt)
          )
          (entmod elst)         ;update block
          (entupd e1)
        )

       )
      )

    )

    (setq attcnt 0          ;setup for next set of attributes
          cnt1   (1+ cnt1)
          e1     e2
          attlst nil
          tmplst nil
          newlst nil
    )
  )
 
)

