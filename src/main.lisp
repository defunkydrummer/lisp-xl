(in-package :lisp-xl)

;; --- NOTE: to be reviewed...
;;(declaim (optimize (speed 1) (debug 3) (safety 3) (space 0)))

;; --- NOTE: to be reviewed as well...  ----
;;(defparameter *gc-every-x-rows* 10000)

(defparameter *element-type* '(unsigned-byte 8))

(defparameter *dummy-vector* (make-array 1 :element-type 'string :adjustable T :fill-pointer 0))
;; Struct for containing sheet information
(defstruct sheet
  (unique-strings *dummy-vector*
                  :type (vector string))
  number-formats
  date-formats 
  file-name ;; temp file name
  original-file-name
  last-stream-position ;;last stream position read. Unused for now.
  )

(defmethod print-object ((s sheet) stream)
  (format stream "#<EXCEL SHEET for ~A, ~D unique strings loaded.>" (sheet-original-file-name s)
          (length (sheet-unique-strings s))))

(defun close-sheet (s)
  "Deletes the in-memory data for the sheet and deletes the temporary file created for it."
  (declare (type sheet s))
  (delete-file (sheet-file-name s))
  (setf (sheet-file-name s) nil)
  (setf (sheet-unique-strings s) *dummy-vector*)
  (setf (sheet-number-formats s) nil)
  (setf (sheet-date-formats s) nil))

(defun read-sheet (file sheet &key (silent nil))
  "Read excel file sheet into new sheet struct.
  'Sheet' must be an index, or sheet name. Nil for first sheet."
  (let* ((sheet-struct (make-sheet))
         (sheets (list-sheets file))
	 (entry-name (cond ((and (null sheet))
			    (third (first sheets))) ;path to first sheet
			   ((stringp sheet)
			    (third (find sheet sheets :key #'second :test #'string=)))
			   ((numberp sheet)
			    (third (find sheet sheets :key #'first))))))
    (unless entry-name
      (error "specify one of the following sheet ids or names: ~{~&~{~S~^~5T~}~}"
	     (loop for (id name) in sheets collect (list id name))))
    (unless silent (format t "Reading file ~A~%" file))
    (zip:with-zipfile (zip file)
      ;; get metadata needed for further process
      (unless silent (format t  "Loading metadata into RAM..."))
      (setf (sheet-unique-strings sheet-struct) (get-unique-strings-vector zip))
      (unless silent (format t  "~D unique strings found.~%"
                             (length (sheet-unique-strings sheet-struct))))
      (setf (sheet-number-formats sheet-struct) (get-number-formats zip))
      (setf (sheet-date-formats sheet-struct)
            (%get-date-formats (sheet-number-formats sheet-struct)))
      ;; read /expand the sheet data itself
      (let ((entry (zip:get-zipfile-entry (format nil "xl/~A" entry-name)
                                          zip)))
        ;; get-zipfile-entry: Return an entry handle for the file called name.
        (when entry
          ;; Load ZIP into temp file
          (let ((temp-file-name
                  (uiop/stream::get-temporary-file :prefix "lisp-xl-temp"
                                                   :suffix ".tmp")))            
            (unless silent
              (format t  "Uncompressing to File [~A] ...~%" temp-file-name))
            (with-open-file (fstream temp-file-name :direction :output
                                                    :if-exists :overwrite
                                                    :element-type *element-type*)
              
              (zip:zipfile-entry-contents entry fstream))
            (setf (sheet-file-name sheet-struct) temp-file-name)
            (setf (sheet-original-file-name sheet-struct) file)
            sheet-struct
            ))))))

(defmacro with-open-excel-sheet ((pathname sheet-index sheet-symbol &optional (silent T)) &body body)
  "Open excel sheet, and execute body. Sheet struct is bound to sheet-symbol"
  (let ((result-var (gensym "result-")))
    `(let* ((,sheet-symbol (read-sheet ,pathname ,sheet-index :silent ,silent))
            (,result-var
              (unwind-protect
                   (progn ,@body))))
       (close-sheet ,sheet-symbol)
       ,result-var)))

;; similar but uses all sheets
(defmacro with-all-excel-sheets ((pathname sheet-symbol index-symbol name-symbol &optional (silent T)) &body body)
  "For every sheet in excel file, execute body. Sheet struct is bound to sheet-symbol.
   Sheet index is bound to index-symbol.
   Sheet name is bound to name-symbol"
  `(loop for ,index-symbol in (list-sheet-indexes ,pathname)
         for ,name-symbol = (sheet-name ,pathname ,index-symbol)
         do
         (with-open-excel-sheet (,pathname ,index-symbol ,sheet-symbol ,silent)
           ,@body)))


;; NOTE: New version, uses klacks instead of cl-xmlspam
;; klacks is part of CXML
(defun %%process-sheet (sheet-struct
                        &key
                          row-begin-function
                          column-process-function
                          row-end-function
                          final-function
                          (max-row nil)
                          (initial-row 1)) 
                                        ;(initial-stream-position nil) 
  
  "Generalized Process sheet (as struct)"
  (declare (type sheet sheet-struct)
           (type fixnum initial-row)
           (type (or null fixnum) max-row)
           (type (function (fixnum (or string null) (or string null)(or string null) (or string null))
                           (or cons null)) 
                 column-process-function)
           (type function row-begin-function
                 row-end-function
                 final-function))
  (let* ((filename (sheet-file-name sheet-struct))
         (source (cxml:make-source filename)))
    (klacks:with-open-source (klacks-source source)
      (let ((style nil)
            (value nil)
            (type nil)
            (position nil)
            (col-cons ())
            (row-cons ())
            (cell-attributes ())
            (cell-value-cons ())
            (col-index 0)
            (row-index 0)
            (row-index-str nil))
        (declare (type fixnum col-index row-index)
                 (type (or cons null) row-cons col-cons cell-attributes cell-value-cons)
                 (type (or string null) style type value position row-index-str))
        
        (block process
          (prog ((builder (cxml-xmls:make-xmls-builder :include-default-values nil
                                                       :include-namespace-uri nil )))
           l_row
             ;; --------- GC at every 50000 rows, what a shame...-----
             ;; to be reviewed! it seems KLACKS is consing too much. 
             #+sbcl(if (eql (mod row-index *gc-every-x-rows*) 0)
                       (sb-ext:gc))
             ;; ------------------------------------------------------
             ;; find the row start
             (handler-case
                 (prog ()
                    (klacks:find-element source "row")
                                        ; serialize the whole row (all columns)
                    (setf row-cons (klacks:serialize-element source builder)))
               (error ()
                 (go l_finish)))        ;like handling EOF, basically
             ;; pop "row" label and "row" attributes, read "row" attributes in particular the row...
             (pop row-cons)
             (setf row-index-str (the (or string null)
                                      (car (cdr (assoc "r" (pop row-cons) :test #'string= )))))
             (when row-index-str
               (setf row-index (parse-integer row-index-str))
               (when max-row
                 (if (> row-index max-row) (go l_finish))))
             ;; skip row if necessary
             (if (< row-index initial-row)
                 (go l_row))
             (funcall row-begin-function row-index)
             ;; iterate over each col-cons, all left is columns
             (setf col-index 0)
             
           l_column
             ;; clear values
             (setf style nil
                   value nil
                   type nil
                   position nil)
             (incf col-index)
             (setf col-cons (cdr (pop row-cons))) ; column info, like  (("t" "s") ("s" "3") ("r" "A6")) ("v" NIL "18"))
             (setf cell-attributes (car col-cons)) ; ("t" "s") ("s" "3") ("r" "A6")
             (setf cell-value-cons (second col-cons)) ; ("v" NIL "18")
             ;; we process the attributes t and s
             (setf style (second (assoc "s" cell-attributes :test #'string=)))
             (setf type  (second (assoc "t" cell-attributes :test #'string=)))
             ;; attribute r has the position (i.e. "A1", "B1" etc)
             (setf position  (second (assoc "r" cell-attributes :test #'string=)))
             ;; now we get the value cons
             (when cell-value-cons ; sometimes there will be no value...
               (setf value (third cell-value-cons)))
             ;; call function for processing all this cell
             (funcall column-process-function row-index value type style position)
             ;; if there are rows in row-cons, we repeat for the next column info
             (if (not (null row-cons)) (go l_column))
           l_rowend
             ;; end of row
             (funcall row-end-function row-index)
             ;; go to next row. EOF should terminate all this.
             (go l_row)
           l_finish
             ;; end of file. Return from block the value obtained by final-function.
             (return-from process (funcall final-function))
             ))
        
        ))))




(defun process-sheet (sheet-struct &key (max-row nil)
                                        (initial-row 1)
                                        (row-function nil)
                                        (silent nil)
                                        (debug-print nil))
                                        ;(initial-stream-position nil)) ;; for future use...
"Process sheet (as struct). Reads rows from the sheet-struct and returns them as cons.
  Important options: 
  * max-row & initial-row control the rows to fetch. First row = 1. 
 
  * row-function receives a lambda where it gets passed, as an argument, the row's column values (as a list).
                 This can be used, for example, for uploading each row to a DB, etc. 
                 If row-function is defined, then <the row values will not be returned by process-sheet.
"
  (declare (type sheet sheet-struct)
           (type fixnum initial-row)
           (type (or fixnum null) max-row)
           (type boolean silent debug-print)
           (type (or (function (cons)) null) row-function))
  (let* ((date-formats (sheet-date-formats sheet-struct))
         (number-formats (sheet-number-formats sheet-struct))
         (unique-strings (sheet-unique-strings sheet-struct))
         ;; pre-process decisions on how to handle cells
         (decision-vector (%get-col-parse-decisions-vector date-formats number-formats))
         (col-cons nil)
         (row-cons nil)
         (last-position 0))
    (declare (type (or cons null) row-cons col-cons date-formats number-formats)
             (type (vector string) unique-strings)
             (type (vector fixnum) decision-vector)
             (type fixnum last-position))           
    (flet 
        
        ((row-begin-function (row-index)
           (declare (type fixnum row-index))
           (setf col-cons nil)
           (setf last-position 0)
           (when debug-print (format t "~%Row Number = ~D || " row-index)))
         (column-process-function (row-index value type style position)
           (declare (ignore row-index)
                    (type (or string null) value type style position))
           (unless position
             (error "No position info for cell."))
           (when debug-print (format t "|~A " position))
           ;; log format ID, type and value
           (when debug-print
             (if number-formats
                 (if style
                     (format t "(number-format: ~D [style=~D],~A,~A) "
                             ;; number format type ID 
                             (car (elt number-formats (parse-integer style)))
                             style
                             type
                             value)
                     (format t "(NULL STYLE,~A,~A) "
                             ;; number format type ID 
                             type
                             value))
                 ;; no number formats
                 (format t "(style ~A, type ~A, value ~A)"
                         style type value)
                 ))
           ;; IMPORTANT:
           ;; if the empty cells are not present in the XML (i.e. we have values for A, B, C, and then H comes next)
           ;; fill the missing column values with NIL
           (unless (zerop last-position)
             (let* ((pos (excel-position-to-column-number position))
                    (diff (- pos last-position)))
               ;; check if we need to fill missing column values
               (when (> diff 1)
                 ;; pad with NIL
                 (dotimes (dummy (1- diff))
                   ;(declare (ignore dummy))
                   (push nil col-cons)))))
           ;; update last-position
           (setf last-position
                 (excel-position-to-column-number position))
           ;; store column value
           (push (%read-col value
                            type
                            style
                            unique-strings
                            decision-vector)
                 col-cons))
         (row-end-function (row-index)
           (declare (type fixnum row-index) (ignore row-index))
           (let ((cc (nreverse col-cons)))
             (declare (type (or null cons) cc))
             (if (null row-function)
                 (push cc row-cons)     ;store row
                 ;; else: use the row-function to process the row
                 (funcall row-function cc) ; pass columns as list. 
                 ))) 
         (final-function ()
           (unless silent (print "Finalizing rows..."))
           (nreverse row-cons)))        ; return rows
    
      ;; call to the function that performs the actual process
      (%%process-sheet sheet-struct
                       :row-begin-function #'row-begin-function
                       :column-process-function #'column-process-function
                       :row-end-function #'row-end-function
                       :final-function #'final-function
                       :max-row max-row
                       :initial-row initial-row))))
                       ;:initial-stream-position initial-stream-position

 


;; simple helper
(defun sheet-first-row (sheet-struct)
  "Obtain first row of sheet"
  (car (process-sheet sheet-struct :max-row 1 :silent T)))


;; helper for column info
(defstruct column-info
  format-id
  format-keyword
  type)

;; (defun report-cells-type-change (sheet-struct &key (max-row nil)
;;                                                    (initial-row 1)
;;                                                    (column-list nil)) ;; get only selected columns (list of indexes)
;;   "Inspect how cell type changes from row to row. Returns number of changes."
;;   (declare (type sheet sheet-struct)
;;            (type fixnum initial-row)
;;            (type (or null fixnum) max-row)
;;            (type (or cons null) column-list))
;;   (let* ((number-formats (sheet-number-formats sheet-struct))
;;          (info nil)
;;          (number-of-changes 0)
;;          (blank-row T))
;;     ;; (col-count 0))
;;     (declare (type (or cons null) info number-formats)
;;              (type (boolean) blank-row)    
;;              (type fixnum number-of-changes))           
;;     (flet ((row-begin-function (row-index)
;;              (declare (type fixnum row-index) (ignore row-index))
;;              (setf blank-row T))        ; reset blank value (row) flag
                                       
;;            (column-process-function (row-index col-index value type style)
;;              (declare (type fixnum row-index col-index)
;;                       (type (or string null) value type style ))
;;              ;; compare style, type  for this column, versus previous.
        
;;              ;; detect blank values - set to T if all row is blank or blank strings
;;              (setf blank-row (and blank-row
;;                                   (or (null value) (eq 0 (length value)))))
;;              (let* ((fmt-id (car (elt number-formats
;;                                       (parse-integer (if (null style) "0" style)))))
;;                     (format-keyword (%get-format-type fmt-id))
;;                     (cinfo (cdr (assoc col-index info)))
;;                     (equal?
;;                       (and cinfo ;; there is info for this column
;;                            (equal (column-info-format-keyword cinfo) format-keyword)
;;                            (equal (column-info-type cinfo) type))))
;;                (declare (type fixnum fmt-id)
;;                         (type (or symbol null) format-keyword)
;;                         (type (or column-info null) cinfo))
;;                (unless equal?
;;                  ;; different style for this column, protest:
;;                  (incf number-of-changes)
;;                  (format t "Row ~D: Column ~D changed to format->~A(~D) | type-> ~A ~%"
;;                          row-index col-index format-keyword fmt-id type)
;;                  ;; store new style info for this column and remove old.
;;                  (setf info
;;                        (cons (cons col-index (make-column-info :type type
;;                                                                :format-id fmt-id
;;                                                                :format-keyword format-keyword))
;;                              (remove (assoc col-index info) info))))))
            
;;            (row-end-function (row-index)
;;              (declare (type fixnum row-index))
;;              (when blank-row
;;                (format t "~%*********** WARNING : row ~D all blank. ************ ~%~%" row-index)
;;                ))
;;            (final-function  ()
;;              (format t "Number of changes: ~D" number-of-changes)
;;              number-of-changes
;;              ))
;;       ;; call to the function that performs the actual process
;;       (if (sheet-number-formats sheet-struct)
;;           (%%process-sheet sheet-struct
;;                            :row-begin-function #'row-begin-function
;;                            :column-process-function #'column-process-function
;;                            :row-end-function #'row-end-function
;;                            :final-function #'final-function
;;                            :max-row max-row
;;                            :initial-row initial-row
;;                            :column-list column-list)
;;           (error "No number formats information available for sheet ~A" sheet-struct)))))


;; (defparameter *safety-stream-position-margin* 100
;;   "Number of bytes to 'rewind' the stream position before continuing reading, for safety")
;; (defparameter *minimum-position-margin* (* 10 *safety-stream-position-margin*)
;;   "Number of bytes to read until we dare to move the initial-stream-position on reads.")

;; -- needs review, because the last-stream-position returned by our XML reading library
;; -- is useless! 
;; (defun factory-batch-fetch (batch-size  ;in rows
;;                             sheet)      ;struct
;;   "Returns closure for fetching the sheet in batches (i.e. 5 rows at a time).
;;   Invoke (funcall) the closure with no arguments. "
;;   (declare (type fixnum batch-size)
;;            (type sheet sheet))
;;   (let ((next-row 1)
;;         (last-stream-position 0)
;;         (result nil))
;;     (declare (type fixnum next-row)
;;              (type (or null cons) result))
;;     (flet ((perform-batch ()
;;              (setf result (process-sheet sheet :initial-row next-row
;;                                                :max-row (+ next-row (- batch-size 1))
;;                                                :silent T
;;                                                :initial-stream-position
;;                                          ;; compute initial stream pos. for reading the data
;;                                          (if (>= last-stream-position *minimum-position-margin* )
;;                                              (- last-stream-position *safety-stream-position-margin*)
;;                                              ;; else
;;                                              0))) ; read from start.

;;              (setf last-stream-position (sheet-last-stream-position sheet))
;;              (setf next-row (+ next-row batch-size))
;;              result 
;;              ))
;;       #'perform-batch ;; return the function 
;;       )))
