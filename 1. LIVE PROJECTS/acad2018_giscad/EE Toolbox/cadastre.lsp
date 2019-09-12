(defun CADASTREP (lay /)
(setq vp (ssget "x" (list '(0 . "viewport")(cons 69 (getvar "cvport"))))
      vpdat (entget (ssname vp 0) '("ACAD")) 
      vpdat (cdadr (assoc -3 vpdat))
      llist nil
);;setq

(foreach n vpdat (if (eq 1003 (car n))(setq llist (append llist (list (cdr n))))))

(if (member lay llist)
 (command "VPLAYER" "T" "*Cadastre*" "C" "")
 (command "VPLAYER" "F" "*Cadastre*" "C" "")
);if

);;defun

(defun CADASTREM (lay /)
(if (eq (cdr (assoc 70 (tblsearch "LAYER" lay))) 1)
 (command "-LAYER" "T" "*Cadastre*" "")
 (command "-LAYER" "F" "*Cadastre*" "")
);if
);;defun

(defun C:TCADASTRE (/ vp vpdat exdat llist cadlays cadlay)
(setq rep 0)
(setq cadlays (list "Cadastre Buildings_SP_EXT" "Cadastre Developer_SP_EXT"
                    "Cadastre Developer_SP_NEW" "Cadastre Easements_SP_EXT"
                    "Cadastre Land Parcels_SP_EXT" "Cadastre Othew Property_SP_EXT"
                    "Cadastre Road Names_SP_HV_LV_SL_EXT"
                    "Cadastre Roads_SP_EXT" "Cadastre Roads_SP_NEW"
                    "Cadastre Suburbs_SP_EXT" "Cadastre Water_SP_EXT"))

(repeat (length cadlays)
 (if (tblsearch "LAYER" (nth rep cadlays))
  (setq cadlay (cdr (assoc 2 (tblsearch "LAYER" (nth rep cadlays)))))
 )
 (setq rep (1+ rep))
)

(if cadlay
 (if (eq (getvar "TILEMODE") 1)
  (CADASTREM cadlay)
  (CADASTREP cadlay)
 );if
 (princ "***No Cadastre Layers Pesent in Drawing***")
);if
(princ)
)