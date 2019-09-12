;Make a 'Clean' Wblock of only what is selected.  (Not suitable for Mspace/Pspace drawings) Discard changes!
;
;	Author:
;		Henry C. Francis
;		425 N. Ashe St.
;		Southern Pines, NC 28387
;
;	http://www.pinehurst.net/~pfrancis
;	e-mail hfrancis@pinehurst.net
;	All rights reserved.
;
(defun c:wbcln ()
  (setq exprt (getvar"expert"))
  (setvar "expert" 3)
  (princ "\nSelect objects for Clean Wblock: ")
  (setq ss (ssget))
  (command ".wblock" (getvar ss) "" (getvar"insbase") ss "")
  (setvar "expert" exprt)
  (if align_lst
    (progn
      (setq num (load_dialog "gpdgn"))
      (gate_dlg)
    )
  )
  (command ".open" "y")
;  (command ".open" "y" "~")
);defun
