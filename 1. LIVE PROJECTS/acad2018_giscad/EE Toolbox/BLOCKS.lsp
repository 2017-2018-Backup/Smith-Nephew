;;;--- BLOCKS.lsp   -  Retrieve quanties of blocks filtering by block name, tag, and or value.
;;;
;;;
;;;--- Copyright 2005 by JefferyPSanders.com
;;;    All rights reserved.
;;;
;;;
;;;--- Revised 8/15/05 - Added the ability to count blocks without attributes.







;;;--- Function to get the block data

(defun getAttData(/ eset hdrStr dataList blkCntr en enlist blkType blkName entName 
                    entPoint entRot entX entY entZ entLay attTag attVal entSty entClr
                    dStr group66)
  ;;;--- Set up an empty list
  (setq dataList(list))

  ;;;--- If that type of entity exist in drawing
  (if (setq eset(ssget "X" (list (cons 0 "INSERT"))))
    (progn

      ;;;--- Set up some counters
      (setq blkCntr 0 cntr 0)
                     
      ;;;--- Loop through each entity
      (while (< blkCntr (sslength eset))
                                
        ;;;--- Get the entity's name
        (setq en(ssname eset blkCntr))
                         
        ;;;--- Get the DXF group codes of the entity
        (setq enlist(entget en))

        ;;;--- Get the name of the block
        (setq blkName(cdr(assoc 2 enlist)))
                      
        ;;;--- Check to see if the block's attribute flag is set
        (if(cdr(assoc 66 enlist))
          (progn
                                   
            ;;;--- Get the entity name
            (setq en(entnext en))
                   
            ;;;--- Get the entity dxf group codes
            (setq enlist(entget en))
                    
            ;;;--- Get the type of block
            (setq blkType (cdr(assoc 0 enlist)))
                    
            ;;;--- If group 66 then there are attributes nested inside this block
            (setq group66(cdr(assoc 66 enlist)))
                   
            ;;;--- Loop while the type is an attribute or a nested attribute exist
            (while(or (= blkType "ATTRIB")(= group66 1))
                    
              ;;;--- Get the block type 
              (setq blkType (cdr(assoc 0 enlist)))
                          
              ;;;--- Get the block name         
              (setq entName (cdr(assoc 2 enlist)))
                       
              ;;;--- Check to see if this is an attribute or a block
              (if(= blkType "ATTRIB")
                (progn
                        
                  ;;;--- Save the name of the attribute
                  (setq attTag(cdr(assoc 2 enlist)))
                      
                  ;;;--- Get the value of the attribute
                  (setq attVal(cdr(assoc 1 enlist))) 
                        
                  ;;;--- Save the data gathered
                  (setq dataList(append dataList(list (list blkName attTag attVal))))
                             
                  ;;;--- Increment the counter
                  (setq cntr (+ cntr 1))
                       
                  ;;;--- Get the next sub-entity or nested entity as you will
                  (setq en(entnext en))
                       
                  ;;;--- Get the dxf group codes of the next sub-entity
                  (setq enlist(entget en))
                      
                  ;;;--- Get the block type of the next sub-entity
                  (setq blkType (cdr(assoc 0 enlist)))
                        
                  ;;;--- See if the dxf group code 66 exist.  if so, there are more nested attributes
                  (setq group66(cdr(assoc 66 enlist)))
                ) 
              )
            )
          )
           
          ;;;--- Else, the block does not contain attributes
          (progn
             
            ;;;--- Setup a bogus tag and value
            (setq attTag "" attVal "")

            ;;;--- Save the data gathered
            (setq dataList(append dataList(list (list blkName attTag attVal))))
          )
        )
        (setq blkCntr (+ blkCntr 1))
      )
    )                    
  )
  dataList
)




;;;--- Function to count occurences within a list

(defun goCountData(ele lst / cnt a)
  (setq cnt 0)
  (foreach a lst
    (if(equal a ele)
      (setq cnt(+ cnt 1))
    )
  )
  cnt
)



;;;--- Function to update the data

(defun updateData()

  ;;;--- Setup a list to hold the selected items
  (setq blkDataList(list))

  ;;;--- Save the list setting
  (setq readlist(get_tile "blklist")) 

  ;;;--- Setup a variable to run through the list
  (setq count 1)

  ;;;--- cycle through the list getting all of the selected items
  (while (setq item (read readlist))
    (setq blkDataList(append blkDataList (list (nth item blkList))))
    (while 
      (and
        (/= " " (substr readlist count 1))
        (/= ""   (substr readlist count 1))
      )
      (setq count (1+ count))
    )
    (setq readlist (substr readlist count))
  )


  ;;;--- Setup a list to hold the selected items
  (setq tagDataList(list))

  ;;;--- Save the list setting
  (setq readlist(get_tile "taglist")) 

  ;;;--- Setup a variable to run through the list
  (setq count 1)

  ;;;--- cycle through the list getting all of the selected items
  (while (setq item (read readlist))
    (setq tagDataList(append tagDataList (list (nth item tagList))))
    (while 
      (and
        (/= " " (substr readlist count 1))
        (/= ""   (substr readlist count 1))
      )
      (setq count (1+ count))
    )
    (setq readlist (substr readlist count))
  )


  ;;;--- Setup a list to hold the selected items
  (setq valDataList(list))

  ;;;--- Save the list setting
  (setq readlist(get_tile "vallist")) 

  ;;;--- Setup a variable to run through the list
  (setq count 1)

  ;;;--- cycle through the list getting all of the selected items
  (while (setq item (read readlist))
    (setq valDataList(append valDataList (list (nth item valList))))
    (while 
      (and
        (/= " " (substr readlist count 1))
        (/= ""   (substr readlist count 1))
      )
      (setq count (1+ count))
    )
    (setq readlist (substr readlist count))
  )

  ;;;--- Set up a list to hold all data that matches the block name
  (setq myData(list))
  
  ;;;--- Cycle through all data and find block name matches
  (foreach a dataList
    (if(or(member "All" blkDataList)(member (car a) blkDataList))
       (setq myData(append myData (list a)))
    )
  )

  ;;;--- Set up a list to hold the data that matches the tag name
  (setq myData2(list))

  ;;;--- Filter through all previous matches to find tag matches
  (foreach a myData
    (if(or(member "All" tagDataList)(member (cadr a) tagDataList))
       (setq myData2(append myData2 (list a)))
    )
  )     

  ;;;--- Filter the previous matches for attribute value matches
  (setq myData3(list))
  (foreach a myData2
    (if(or(member "All" valDataList)(member (caddr a) valDataList))
       (setq myData3(append myData3 (list a)))
    )
  )  

  ;;;--- Set up a list to hold the used data [old data]
  (setq usedData(list))

  ;;;--- Reset the orignal list to hold the new data
  (setq myData(list))

  ;;;--- Cycle through each match
  (foreach a myData3

    ;;;--- Make sure we haven't counted it already
    (if(not(member a usedData))
      (progn

        ;;;--- Count the number of times the match appears in the list
        (setq cnt(goCountData a myData3))

        ;;;--- Add the data and the count to the new list         
        (setq myData(append myData (list(list cnt a))))

        ;;;--- Add the data to the used list so we don't count it twice
        (setq usedData(append usedData (list a)))
      )
    )
  )

  ;;;--- Set up a list to hold the formatted data for the dialog box
  (setq newData(list))

  ;;;--- If new data was found
  (if myData
    (progn

      ;;;--- Format the data
      (foreach a myData
        (setq qty(substr (strcat (itoa (car a))"        ") 1 6))
        (setq bna(substr (strcat (car  (cadr a)) "                             ") 1 20))
        (setq tgn(substr (strcat (cadr (cadr a)) "                             ") 1 20))
        (setq van(substr (strcat (caddr(cadr a)) "                             ") 1 20))

        ;;;--- Add it to the new data list
        (setq newData(append newData (list (strcat qty bna tgn van))))
      )
       
      ;;;--- Add a horizontal line
      (setq newData
        (append 
          (list
            "----------------------------------------------------------------------------------------"
          )
          newData
        )
      )

      ;;;--- Add a header
      (setq newData
        (append 
          (list
            "Qty.  Block Name          Tag                 Value"
          )
          newData
        )
      )
    
      ;;;--- Add the list to the dialog box
      (start_list "totallist" 3)
      (mapcar 'add_list newData)
      (end_list)
    )
    (progn
          
      ;;;--- Clear the list in the dialog box
      (start_list "totallist" 3)
      (mapcar 'add_list newData)
      (end_list)
    )        
  )
)




;;;--- Function to sort a list
;;;
;;;--- Usage  (sort (list "F" "A" "B"))
;;;
(defun sort(alist / n)(setq lcup nil rcup nil)
 (defun cts(a b)
  (cond
   ((> a b)t)
   ((= a b )t)
   (t nil)
 ))
 (foreach n alist
  (while (and rcup(cts n(car rcup)))(setq lcup(cons(car rcup)lcup)rcup(cdr rcup)))
   (while (and lcup(cts(car lcup)n))(setq rcup(cons(car lcup)rcup)lcup(cdr lcup)))
   (setq rcup(cons n rcup))
 )
 (append(reverse lcup)rcup)
)



;;;--- Function to write the data to a file

(defun writeData()

  ;;;--- Set up a counter 
  (setq lineCntr 0)

  ;;;--- Get a file name from the user
  (if(setq filName(getfiled "Select File Name" "" "csv" 1))
    (progn 
                   
      ;;;--- Open the file to write
      (if(setq fil(open filName "w"))
        (progn
                    
          ;;;--- Print a header string
          (princ "QTY.,Block Name,Tag,Value\n" fil)
                     
          (foreach a myData
            (princ (strcat (itoa(car a)) "," (car(cadr a))","(cadr(cadr a))","(caddr(cadr a))"\n") fil)
            (setq lineCntr(+ lineCntr 1))
          )

          ;;;--- Close the file
          (close fil)
                                      
          (alert 
            (strcat "Finished Sending " (itoa lineCntr) 
                    " lines to the CSV file!\n\nNote: Double click the CSV file and Excel will open it."
            )
          )
        )
        (princ "\n ERROR - Could not open CSV file!")
      )
    )
    (princ "\n ERROR - Invalid file name!")
  )
)





;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;                                                                          ;;;
;;;     888       888          888           8888888      8888   888         ;;;
;;;     8888     8888         88888            888        88888  888         ;;;
;;;     88888   88888        888 888           888        888888 888         ;;;
;;;     888888 888888       888   888          888        888 888888         ;;;
;;;     888 88888 888      88888888888         888        888  88888         ;;;
;;;     888  888  888     888       888      8888888      888   8888         ;;;
;;;                                                                          ;;;
;;;                                                                          ;;;
;;;                888            888888888        888888888                 ;;;
;;;               88888           888   888        888   888                 ;;;
;;;              888 888          888   888        888   888                 ;;;
;;;             888   888         888888888        888888888                 ;;;
;;;            88888888888        888              888                       ;;;
;;;           888       888       888              888                       ;;;
;;;                                                                          ;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;



(defun C:BLOCKS()

  ;;;--- Setup an empty list to hold all of the data
  (setq myData(list))

  ;;;--- Get a the data from the blocks
  (if(setq dataList(getattData))
    (progn
      
      ;;;--- Setup list to hold the block names, tag names, and values         
      (setq blkList(list) tagList(list) valList(list))

      ;;;--- Seperate the data
      (foreach a dataList
        (if(not (member (car a) blkList))
          (setq blkList(append blkList (list (car a))))
        )
        (if(not (member (cadr a) tagList))
          (setq tagList(append tagList (list (cadr a))))
        )
        (if(not (member (caddr a) valList))
          (setq valList(append valList (list (caddr a))))
        )
      )

      ;;;--- Sort the list
      (setq blkList(sort blkList))
      (setq tagList(sort tagList))
      (setq valList(sort valList))

      ;;;--- Add the "ALL" option
      (setq blkList(append (list "All") blkList))
      (setq tagList(append (list "All") tagList))
      (setq valList(append (list "All") valList))

      ;;;--- Build the data list
      (setq blkDataList blkList tagDataList tagList valDataList valList)
 
      ;;;--- Put up the dialog box
      (setq dcl_id (load_dialog "BLOCKS.dcl"))
 
      ;;;--- See if it is already loaded
      (if (not (new_dialog "BLOCKS" dcl_id))
        (progn
          (alert "The BLOCKS.DCL file could not be found!")
          (exit)
        )
      )


      ;;;--- Add the list to the dialog box
      (start_list "blklist" 3)
      (mapcar 'add_list blkList)
      (end_list)
      (start_list "taglist" 3)
      (mapcar 'add_list tagList)
      (end_list)
      (start_list "vallist" 3)
      (mapcar 'add_list valList)
      (end_list)

      (updateData)

      ;;;--- If an action event occurs, do this function
      (action_tile "blklist" "(updateData)")
      (action_tile "taglist" "(updateData)")
      (action_tile "vallist" "(updateData)")
      (action_tile "write"   "(writeData)")
      (action_tile "cancel"  "(done_dialog)")

      ;;;--- Display the dialog box
      (start_dialog)
    )
    (alert "No blocks with attributes found in drawing!")
  )
  (princ)
)
   