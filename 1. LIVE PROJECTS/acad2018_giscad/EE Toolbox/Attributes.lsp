;;	A series of routines for changing attribute properties on a mass of attributes.
;;	Can change each proprty, i.e. the layer, color, style, layer etc of many attributes at once.
;;
;;	Also a routine to change attribute text to upper case.
;;
;;	Another to transfer data from one attribute to another. Handy if you are updating or
;;	replacing blocks with a new version, i.e. line number racetrack blocks. Also copies from
;;	text to attribute but not from attribute to text.
;;
;;	Another routine if you drawing piping GA's using line numbers for layer names. Get the
;;	layer name from an object and paste into the attribute in a line number racetrack block.
;;


;;;Attribute style change
(defun c:asc ()
(setq atts(getstring "\nEnter new attribute style: "))
(command "attedit" "y" "*" "*" "*" "c" pause pause)
(while
(= 1 (logand (getvar "cmdactive") 1))
(command "s" atts "n")
);;; end while
);;; end lisp

;;;Attribute height change
(defun c:ahc ()
(setq atth(getreal "\nEnter new attribute height: "))
(command "attedit" "y" "*" "*" "*" "c" pause pause)
(while
(= 1 (logand (getvar "cmdactive") 1))
(command "h" atth "n")
);;; end while
);;; end lisp

;;;Attribute angle change
(defun c:aac ()
(setq atta(getint "\nEnter new attribute angle: "))
(command "attedit" "y" "*" "*" "*" "c" pause pause)
(while
(= 1 (logand (getvar "cmdactive") 1))
(command "a" atta "n")
);;; end while
);;; end lisp

;;;Attribute color change
(defun c:acc ()
(setq attc(getint "\nEnter new attribute color: "))
(command "attedit" "y" "*" "*" "*" "c" pause pause)
(while
(= 1 (logand (getvar "cmdactive") 1))
(command "c" attc "n")
);;; end while
);;; end lisp

;;;Attribute layer change
(defun c:alc ()
(setq attl(getstring "\nEnter new attribute layer: "))
(command "attedit" "y" "*" "ASSET_NUM" "*" "c" pause pause)
(while
(= 1 (logand (getvar "cmdactive") 1))
(command "l" attl "n")
);;; end while
);;; end lisp

;;; Block attribute case change
(defun c:AttUpper ()
 (setq ATWC1(nentsel "\nSelect block data from: "))
  (setq ATWC2(car ATWC1))
   (setq ATWC3(entget ATWC2))
    (setq ATWC4(cdr(assoc 1 ATWC3)))
     (setq ATWC5(cdr(assoc 10 ATWC3)))
  (setq ATWCTXT (Strcase ATWC4))
(command "attedit" "n" "y" "*" "*" "*" ATWC5 "" ATWC4 ATWCTXT)
)

;;; Block attribute data transfer
(defun c:BADT ()
 (setq e1s(nentsel "\nSelect block data from: "))
 (setq e2s(nentsel "\nSelect block data to: "))
  (setq e1n(car e1s))
  (setq e2n(car e2s))
   (setq e1d(entget e1n))
   (setq e2d(entget e2n))
    (setq e1dd(cdr(assoc 1 e1d)))
    (setq e2dd(cdr(assoc 1 e2d)))
     (setq e1p(cdr(assoc 10 e1d)))
     (setq e2p(cdr(assoc 10 e2d)))
  (setq tt1 e1dd)
  (setq tt2 e2dd)
(command "attedit" "n" "y" "*" "*" "*" e2p "" tt2 tt1)
)

;;; Layer name to block attribute
(defun c:LTATT ()
 (setq en1s(nentsel "\nSelect entity on layer: "))
 (setq en2s(nentsel "\nSelect block data to: "))
  (setq en1n(car en1s))
  (setq en2n(car en2s))
   (setq en1d(entget en1n))
   (setq en2d(entget en2n))
    (setq en1dd(cdr(assoc 8 en1d)))
    (setq en2dd(cdr(assoc 1 en2d)))
     (setq en1p(cdr(assoc 10 en1d)))
     (setq en2p(cdr(assoc 10 en2d)))
  (setq ttn1 en1dd)
  (setq ttn2 en2dd)
(command "attedit" "n" "y" "*" "*" "*" en2p "" ttn2 ttn1)
)