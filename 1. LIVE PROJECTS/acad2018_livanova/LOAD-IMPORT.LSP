;;;==============================
;;;       LISP EXAMPLE
;;;==============================
(defun C:LOAD-IMPORT (/ shell)
  (vl-load-com)
  (setq filename "C:\\ACAD2018_LivaNova\\Interface\\SSDL-EXACT-IMPORT.exe")
  (setq shell
  (vla-getinterfaceobject
    (vlax-get-acad-object)
    "Shell.Application"
  )
  )
  (vlax-invoke-method shell 'Open filename)
  (vlax-release-object shell)
)
