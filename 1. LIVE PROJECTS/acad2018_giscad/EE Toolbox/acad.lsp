(setvar "cmdecho" 0)
(princ "\nConfiguring GISCAD 2018 Environment...")
(princ "\nLoading GISCAD 2018 Start-Up Files...")
(LOAD "C:/ACAD2018_GISCAD/EE Toolbox/wbcln.lsp")
(LOAD "C:/ACAD2018_GISCAD/EE Toolbox/utils.lsp")
(LOAD "C:/ACAD2018_GISCAD/EE Toolbox/BLOCKS.lsp")
(LOAD "C:/ACAD2018_GISCAD/EE Toolbox/Attributes.lsp")
(LOAD "C:/ACAD2018_GISCAD/EE Toolbox/attflip.lsp")
(LOAD "C:/ACAD2018_GISCAD/EE Toolbox/atr.lsp")
(LOAD "C:/ACAD2018_GISCAD/EE Toolbox/Batch Draw Order.lsp")
(LOAD "C:/ACAD2018_GISCAD/EE Toolbox/loadpdf.lsp")
;(LOAD "C:/ACAD2018_GISCAD/EE Toolbox/acad.lsp")
(LOAD "C:/ACAD2018_GISCAD/EE Toolbox/cadastre.lsp")
(LOAD "C:/ACAD2018_GISCAD/EE Toolbox/CIRCUITVIEWS.lsp")
;(LOAD "C:/ACAD2018_GISCAD/EE Toolbox/xdText.lsp")
(LOAD "C:/ACAD2018_GISCAD/EE Toolbox/LOAD-A2K-SUPPORT-SITE.LSP")

(command "._NETLOAD" "giscad.dll")
(command "._NETLOAD" "giscadu.dll")
(command "._NETLOAD" "copypastattr.dll")
(command "._NETLOAD" "titleblock.dll")
(command "._NETLOAD" "EEVersion.dll")
 
(princ "\nLoading GISCAD 2018 EE Toolbox...")
(arxload "C:\\ACAD2018_GISCAD\\EE Toolbox\\EE_Toolbox_2018.arx")

(princ "\nLoading GISCAD 2018 Internal Validation...")
(arxload "C:\\ACAD2018_GISCAD\\EE Toolbox\\EE_Validation_2018.arx")

(princ "\nGISCAD 2018 Loaded Successfully...\n")

((lambda (/ ent)
;(while (not (setq ent (car (entsel)))))
(setq ss(ssget "x" (list (cons 0 "insert")(cons 2 "a1_frame_title"))))
   (if ss
     (progn
   (setq ent (ssname ss 0))
(setq ent (vlax-ename->vla-object ent))
(vla-ScaleEntity
ent
(vlax-3d-point 0 0 0 )
1
)))

   (setq ss(ssget "x" (list (cons 0 "insert")(cons 2 "a2_frame_title"))))
    (if ss
      (progn
   (setq ent (ssname ss 0))
(setq ent (vlax-ename->vla-object ent))
(vla-ScaleEntity
ent
(vlax-3d-point 0 0 0 )
1
)))

   (setq ss(ssget "x" (list (cons 0 "insert")(cons 2 "a3_frame_title"))))
    (if ss
      (progn
   (setq ent (ssname ss 0))
(setq ent (vlax-ename->vla-object ent))
(vla-ScaleEntity
ent
(vlax-3d-point 0 0 0 )
1
)))

   (setq ss(ssget "x" (list (cons 0 "insert")(cons 2 "a4_frame_title"))))
    (if ss
      (progn
   (setq ent (ssname ss 0))
(setq ent (vlax-ename->vla-object ent))
(vla-ScaleEntity
ent
(vlax-3d-point 0 0 0 )
1
)))
)
)