
==========================================================
(defun c:attrot(/ blSet attLst errCount oldAng)
	(if(not atrot:rAng)(setq atrot:rAng 0))
	(setq oldAng atrot:rAng
	atrot:rAng
	(getangle
	(strcat "\nSpecify rotation angle <"(angtos atrot:rAng)">: ")))
	(if(not atrot:rAng)(setq atrot:rAng oldAng))
	(princ "<<< Select blocks to rotate attributes >>>")
	(setq errCount 0)
	(if
	(setq blSet(ssget '((0 . "INSERT")(66 . 1))))
	(progn
	(setq blSet(mapcar 'vlax-ename->vla-object
	(vl-remove-if 'listp
	(mapcar 'cadr(ssnamex blSet)))))
	(foreach itm blSet
	(setq attLst
	(vlax-safearray->list
	(vlax-variant-value
	(vla-GetAttributes itm))))
	(foreach att attLst
	(if(vl-catch-all-error-p
	(vl-catch-all-apply
	'vla-put-Rotation(list att atrot:rAng)))
	(setq errCount(1+ ErrCount))
	); end if
	); end foreach
	); end foreach
	); end progn
	(princ ">>> Nothing selected! <<<")
	); end if
	(if(/= 0 errCount)
	(princ
	(strcat "\n>>> "
	(itoa errCount)
	" attributes or blocks were on locked layer! <<< "))
	); end if
	(princ)
); end of c:atrot

(defun c:HF ()
	(setq ents (ssget))(command "chprop" ents "" "LA" "features_HIDE" "") 
)

(defun c:DF ()
	(setq ents (ssget))(command "chprop" ents "" "LA" "features_DEL" "") 
)

(defun c:RB ()
	(c:attrot 0)
)

(defun c:cadON ()
	(setq my_spt 1)
	(if(wcmatch (getvar"clayer") "*Cadastre*")
	(command "-layer" "set" "0" "") )
	(command "-layer" "freeze" "*Cadastre*" "") (princ) 
) ;

defun (defun c:cadOFF ()
	(setq my_spt 0)
	(command "-layer" "thaw" "*Cadastre*" "")
	(command "regenall") (princ) 
) ;

defun (defun c:tCAD () 
(if(not my_spt)(setq my_spt 0)) ;
in case cadON or cadOFF has not be run before (if (zerop my_spt) (c:cadON) (c:cadOFF) ) (princ) ) ;

(defun c:HideON ()
	(setq my_HIDE 1)
	(if(wcmatch (getvar"clayer") "Features_HIDE")
	(command "-layer" "set" "0" "") )
	(command "-layer" "ON" "Features_HIDE" "")
	(princ) 
) ;

defun (defun c:HideOFF ()
	(setq my_HIDE 0)
	(command "-layer" "OFF" "Features_HIDE" "")
	(command "regenall")
	(princ) 
) ;

defun (defun c:tHIDE () 
(if(not my_HIDE)(setq my_HIDE 0)) ;
in case HideON or HideOFF has not be run before (if (zerop my_HIDE) (c:HideON) (c:HideOFF) ) (princ) ) ;


(defun c:DelON ()
	(setq my_DEL 1)
	(if(wcmatch (getvar"clayer") "Features_DEL")
	(command "-layer" "set" "0" "") )
	(command "-layer" "ON" "Features_DEL" "")
	(princ) 
) ;

defun (defun c:DelOFF ()
	(setq my_DEL 0)
	(command "-layer" "OFF" "Features_DELE" "")
	(command "regenall")
	(princ) 
) ;

defun (defun c:tDEL () 
(if(not my_DEL)(setq my_DEL 0)) ;
in case DelON or DelOFF has not be run before (if (zerop my_DEL) (c:DelON) (c:DelOFF) ) (princ) ) ;

Defun (defun C:fw ()
 	(vl-vbarun "W:\\AutoCAD\\GISCAD\\Programs\\IE Utils\\VBA\\IE_Utilities.dvb!ThisDrawing.WipeoutFix")
 	(princ)
 ) ;

