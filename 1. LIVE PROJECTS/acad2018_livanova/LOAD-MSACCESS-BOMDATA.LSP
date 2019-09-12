;;;==============================
;;;       LISP EXAMPLE
;;;==============================
(defun C:LOAD-MSACCESS-BOMDATA (/ shell)
  (vl-load-com)
  (setq filename "C:\\ACAD2018_LivaNova\\BOM Database\\bomdata.mdb")
  (setq shell
  (vla-getinterfaceobject
    (vlax-get-acad-object)
    "Shell.Application"
  )
  )
  (vlax-invoke-method shell 'Open filename)
  (vlax-release-object shell)
)
