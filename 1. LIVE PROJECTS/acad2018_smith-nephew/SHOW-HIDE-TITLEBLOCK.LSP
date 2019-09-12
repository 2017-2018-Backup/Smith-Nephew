(defun c:TB_show ( )
    (RedrawSelectionSet
        (ssget "_X"
            (list
               '(0 . "INSERT")
               '(2 . "MAINPATIENTDETAILS")
                (cons 410 (if (= 1 (getvar 'CVPORT)) (getvar 'CTAB) "Model"))
            )
        )
        1
    )
    (princ)
)

(defun c:TB_hide ( )
    (RedrawSelectionSet
        (ssget "_X"
            (list
               '(0 . "INSERT")
               '(2 . "MAINPATIENTDETAILS")
                (cons 410 (if (= 1 (getvar 'CVPORT)) (getvar 'CTAB) "Model"))
            )
        )
        2
    )
    (princ)
)

(defun RedrawSelectionSet ( ss mode / i )
    (if ss
        (repeat (setq i (sslength ss))
            (redraw (ssname ss (setq i (1- i))) mode)
        )
    )
)