;;http://www.cadtutor.net/corner/2008/january.php

(vl-load-com)

(defun c:BatchDrawOrder ( / sDWGFolder sDWGFiles acadApp docsColl sFile dwgObj modelSpace extDict sortentsTable cnt acObject lstTemp lstObjects cntStep)
  ;; Set the folder to update the drawings in
  (setq sDWGFolder "c:\\test")

  ;; Get the list of DWG files in the folder
  (setq sDWGFiles (vl-directory-files sDWGFolder "*.DWG" 1))

  ;; Get the AutoCAD application object
  (setq acadApp (vlax-get-acad-object))

  ;; Get the Documents collection
  (setq docsColl (vla-get-documents acadApp))

  (foreach sFile sDWGFiles
    (progn
      ;; Open the drawing file
      (setq dwgObj (vla-open docsColl (strcat sDWGFolder "\\" sFile) :vlax-false))

      ;; Get the Model Space object for the Drawing
      (setq modelSpace (vla-get-modelSpace dwgObj))

      ;; Get the Extension Dictionary for Model Space
      (setq extDict (vla-getExtensionDictionary modelSpace))
  
      ;; Check to see if the Sortents Tables exists in the drawing
      (if (= (type (setq sortentsTable (vl-catch-all-apply 'vlax-invoke-method (list extDict 'Item "ACAD_SORTENTS")))) 'VL-CATCH-ALL-APPLY-ERROR)
        (setq sortentsTable (vla-addObject extDict "ACAD_SORTENTS" "AcDbSortentsTable"))
      )

      ;; Set counter for text to zero
      (setq cnt 0 lstObjects nil)

      ;; Loop through the ModelSpace collection
      (vlax-for acObject modelSpace

        ;; Check to see if the object is a text object
        (if (or (= (vla-get-objectName acObject) "AcDbText")(= (vla-get-objectName acObject) "AcDbMText"))
          (progn

  	    ;; Adjust the array based on the number of objects in the drawing
	    (if (>= cnt 1)
	      (progn
	        ;; Get the current values in the array and then adjust the size
	        (setq lstTemp lstObjects)
                (setq lstObjects (vlax-make-safearray vlax-vbObject (cons 0 cnt)))

	        ;; Step through the old array values and load them into the new array
	        (setq cntStep 0)
	        (repeat cnt 
	          (vlax-safearray-put-element lstObjects cntStep (vlax-safearray-get-element lstTemp cntStep))
	          (setq cntStep (1+ cntStep))
	        )
	      )
	  
	      (setq lstObjects (vlax-make-safearray vlax-vbObject '(0 . 0)))
            )

	    ;; Add the text object to the array and increment counter by 1
	    (vlax-safearray-put-element lstObjects cnt acObject)
	    (setq cnt (1+ cnt))
          )
        )
      )

      ;; If the array is not empty then adjust the text objects Draw Order value
      (if (/= lstObjects nil)
	(progn
          (vla-moveToTop sortentsTable lstObjects)

          ;; Save the changes to the drawing
          (vla-save dwgObj)
	)
      )
      
      ;; Close the drawing and go to the next one
      (vla-close dwgObj)
      
    )
  )
  
 (princ)
)