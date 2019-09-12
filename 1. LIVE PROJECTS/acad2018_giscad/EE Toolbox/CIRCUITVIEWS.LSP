(defun CIRCUITVIEWS(str / lay)
(setvar "CMDECHO" 0)
(command "-LAYER" "T" "0" "")
(setvar "CLAYER" "0")
(if (eq (getvar "TILEMODE") 1)
 (command "-LAYER" "F" "*" "")
 (command "VPLAYER" "F" "*" "" "")
);if
(setq lay (cdr (assoc 2 (tblnext "LAYER" T))))
(while lay
 (if (or (vl-string-search (strcase str) lay)(vl-string-search "Cadastre" lay))
  (if (eq (getvar "TILEMODE") 1)
   (command "-LAYER" "T" lay "")
   (command "VPLAYER" "T" lay "" "")
  );if
 );if
 (setq lay (cdr (assoc 2 (tblnext "LAYER"))))
);while
);defun
