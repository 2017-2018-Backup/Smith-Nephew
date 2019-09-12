;;;==============================
;;;       LISP EXAMPLE
;;;==============================
(defun C:LOAD-A2K-SUPPORT-SITE (/ shell)
  (vl-load-com)
  (setq filename "http://support.a2ktechnologies.com.au/")
  (setq shell
  (vla-getinterfaceobject
    (vlax-get-acad-object)
    "Shell.Application"
  )
  )
  (vlax-invoke-method shell 'Open filename)
  (vlax-release-object shell)
)
