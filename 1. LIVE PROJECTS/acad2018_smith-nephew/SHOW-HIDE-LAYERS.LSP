(defun C:LAYER_ON ()
  (command
    "_.layer"
      "_thaw" "*"
      "_set" "1" ; to ensure the one left un-Frozen is one you want on at end
      "_on" "DATA,CONSTRUCT"
      "" ; complete Layer command
    "_regen" ; in case any of them were Frozen before, and require Regen to show again
  ); command
  (princ)
); defun

(defun C:LAYER_OFF ()
  (command
    "_.layer"
      "_thaw" "*"
      "_set" "1" ; to ensure the one left un-Frozen is one you want on at end
      "_off"  "DATA,CONSTRUCT"
      "" ; complete Layer command
    "_regen" ; in case any of them were Frozen before, and require Regen to show again
  ); command
  (princ)
); defun

(defun C:NOTES_ON ()
  (command
    "_.layer"
      "_thaw" "*"
      "_set" "1" ; to ensure the one left un-Frozen is one you want on at end
      "_on"  "NOTES"
      "" ; complete Layer command
    "_regen" ; in case any of them were Frozen before, and require Regen to show again
  ); command
  (princ)
); defun

(defun C:NOTES_OFF ()
  (command
    "_.layer"
      "_thaw" "*"
      "_set" "1" ; to ensure the one left un-Frozen is one you want on at end
      "_off"  "NOTES"
      "" ; complete Layer command
    "_regen" ; in case any of them were Frozen before, and require Regen to show again
  ); command
  (princ)
); defun

(defun C:ALL_ON ()
  (command
    "_.layer"
      "_thaw" "*"
      "_set" "1" ; to ensure the one left un-Frozen is one you want on at end
      "_on"  "*"
      "" ; complete Layer command
    "_regen" ; in case any of them were Frozen before, and require Regen to show again
  ); command
  (princ)
); defun