(defun c:UGlayers (/ lay rep llst fil)
(setvar "CMDECHO" 0)
(command "-LAYER" "T" "0" "")
(setvar "CLAYER" "0")
(setq rep 0 llst nil)
(setq fil (open (findfile "UGlayers.txt") "r"))

(while (setq lay (read-line fil))
 (setq llst (append llst (list lay)))
);while
(close fil)

(if (eq (getvar "TILEMODE") 1)
 (command "-LAYER" "F" "*" "")
 (command "VPLAYER" "F" "*" "" "")
);if

(setq lay (cdr (assoc 2 (tblnext "LAYER" T))))
(while lay
 (if (vl-string-search "Cadastre" lay)
  (if (eq (getvar "TILEMODE") 1)
   (command "-LAYER" "T" lay "")
   (command "VPLAYER" "T" lay "" "")
  );if
 );if
 (setq lay (cdr (assoc 2 (tblnext "LAYER"))))
);while

(repeat (length llst)
 (setq lay (nth rep llst))
 (if (tblsearch "LAYER" lay)
  (if (eq (getvar "TILEMODE") 1)
   (command "-LAYER" "T" lay "")
   (command "VPLAYER" "T" lay "" "")
  );if
 );if
 (setq rep (+ rep 1))
);repeat
(princ)
);defun
