;;;==============================
;;;       LISP EXAMPLE
;;;==============================
(defun C:LaunchPDF (/ shell)
  (vl-load-com)
  (setq filename "C:\\ACAD2018_GISCAD\\HELP\\GISCAD 2018 USER GUIDE.pdf")
  (setq shell
  (vla-getinterfaceobject
    (vlax-get-acad-object)
    "Shell.Application"
  )
  )
  (vlax-invoke-method shell 'Open filename)
  (vlax-release-object shell)
)
