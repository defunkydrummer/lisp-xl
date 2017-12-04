(in-package :cl-user)
(defpackage :lisp-xl
  (:use #:cl #|:xspam|# :lisp-xl.metadata)
  (:export #:list-sheets
           #:read-sheet
           #:process-sheet
           #:close-sheet
           #:sheet-first-row
           #:with-open-excel-sheet
           #:reports-cells-type-change
   ))
(in-package :lisp-xl)

;; --- NOTE: to be reviewed...
(declaim (optimize (speed 3) (debug 0) (safety 0)))

;; From Carlos Ungil
(defun %get-date-formats (number-formats)
  "Filter (obtain) which formats are Date formats, from the number-formats list"
  (loop for f in number-formats
        for id = (the fixnum (car f))
        for fmt = (cdr f)
        for is-date = 
        (or (<= 14 id 17) ;; built-in: m/d/yyyy d-mmm-yy d-mmm mmm-yy
            (and (stringp fmt) (not (search "h" fmt)) (not (search "s" fmt))
                 (search "d" fmt) (search "m" fmt) (search "y" fmt)))
        collect is-date))

;; From Carlos Ungil
(defun %excel-date (int)
  "Decode the excel date values into something we can parse."
  (declare (type fixnum int))
  (the string 
       (apply #'format nil "~D-~2,'0D-~2,'0D"
              (reverse (subseq (multiple-value-list (decode-universal-time (* 24 60 60 (- int 2)))) 3 6)))))

;; valid format types
(defparameter *formats*
  '(:string :int :real :percent :scientific :time :date :datetime :custom))

(defparameter *number-formats*
  '(:int :real :percent))

(defparameter *unsupported-formats*
  '(:scientific :time :datetime :unsupported :custom))

(declaim (type cons *formats* *number-formats* *unsupported-formats*))

(defun %get-format-type (id)
  "Obtain format type (keyword) according to the ID of the format, according to the XLSX standard."
  (declare (type fixnum id))
  (the symbol 
       (cond
         ((eq id 0) :string)
         ((eql id 1) :int)
         ((or (<= 2 id 8) (<= 37 id 40)) :real)
         ((<= 9 id 10) :percent)
         ((or (eql id 11) (eql id 48)) :scientific)
         ((<= 14 id 17) :date)
         ((<= 18 id 21) :time)
         ((eql id 22) :datetime)
         ((>= id 164) :custom) ;; custom formats (found on number-formats list)
         (t :unsupported))))
      
(defun %format-numeric-p (id)
  "Check if Format ID is of a numeric type according to the XLSX standard."
  (declare (type fixnum id))
  (the boolean 
       (when (member (%get-format-type id)
                     *number-formats*) t)))

;; decisions on how to parse cell values
(defparameter *decision-parse-string* 0)
(defparameter *decision-parse-number* 1)
(defparameter *decision-try-lisp-reader* 2)
(defparameter *expected-number-formats-max-size* 512)

(defun %get-col-parse-decisions-vector (date-formats number-formats)
  (declare (ignore date-formats))
  "Inner function. Obtain vector of decisions on how to parse the column value according to the 'style' ID"
  (let ((vector (make-array *expected-number-formats-max-size*
                            :element-type 'fixnum :adjustable nil)))
    (loop with vector-index = 0
          for nf in number-formats
          for format-id = (car nf)
          for format-type = (%get-format-type format-id)
       
          for decision = (cond
                           ((equal format-type :string) *decision-parse-string*)
                           ((member format-type *number-formats*) *decision-parse-number*)
                           ((member format-type *unsupported-formats*) *decision-try-lisp-reader*)
                           (t *decision-try-lisp-reader*)) ; note! 
          do (setf (aref vector vector-index) decision)
             (incf vector-index))
    vector
    ))
  

(defun %read-col (value type style unique-strings decision-vector)
  "Inner function for reading(parsing) a column"
  (declare (type (or string null) value style type)
           (type (vector string) unique-strings)
           (type (simple-array fixnum) decision-vector))
  (let* ((nstyle (the fixnum (parse-integer style)))
         (decision (aref decision-vector nstyle)))
    (declare (type fixnum nstyle decision))
    (handler-case 
        (cond ((null value) nil) ; value is nil (case for blank cells)
              ;; ((string= type "e") (intern value "KEYWORD"))
              ((not (null type))
               (cond 
                 ((string= type "str") value) ;; CALCULATED STRING, we just return the value
                 ;; string from the 'unique-strings' file?
                 ((string= type "s")
                  (if
                   ;; number 
                   (eq decision *decision-parse-number*) (read-from-string value)
                   ;; else - read string from table
                   (aref unique-strings (the fixnum (parse-integer value)))))))
              ;; date, etc
              ;; for now, we try the lisp reader
              (t (read-from-string value))) 
      (error () "PARSE ERROR") ; return "ERROR" if parse errors were found
      )))

(defparameter *element-type* '(unsigned-byte 8))

(defparameter *dummy-vector* (make-array 1 :element-type 'string :adjustable T :fill-pointer 0))
;; Struct for containing sheet information
(defstruct sheet
  (unique-strings *dummy-vector*
                  :type (vector string))
  number-formats
  date-formats 
  file-name ;; temp file name
  last-stream-position ;;last stream position read. Unused for now.
  )

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
	 (entry-name (cond ((and (null sheet) (= 1 (length sheets)))
			    (third (first sheets)))
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
            sheet-struct
            ))))))

(defmacro with-open-excel-sheet ((pathname sheet-index sheet-symbol &optional (silent T)) &body body)
  "Open excel sheet, and execute body. Sheet struct is bound to sheet-symbol"
  `(let ((,sheet-symbol (read-sheet ,pathname ,sheet-index :silent ,silent)))
      (unwind-protect (progn ,@body))
      (close-sheet ,sheet-symbol)))


;; NOTE: New version, uses klacks instead of cl-xmlspam
;; klacks is part of CXML
(defun %%process-sheet (sheet-struct
                         &key
                           row-begin-function
                           column-process-function
                           row-end-function
                           final-function
                           (max-row nil)
                           (initial-row 1)
                           (column-list nil))  ; get only selected columns (list of indexes)
                                        ;(initial-stream-position nil) 
  
  "Generalized Process sheet (as struct)"
  (declare (type sheet sheet-struct)
           (type fixnum initial-row max-row)
           (type function row-begin-function
                 column-process-function
                 row-end-function
                 final-function)
           (type (or cons null) column-list))
  (let* ((filename (sheet-file-name sheet-struct))
         (source (cxml:make-source filename)))
    (klacks:with-open-source (klacks-source source)
      (let ((style nil)
            (value nil)
            (type nil)
            (col-cons ())
            (row-cons ())
            (cell-attributes ())
            (cell-value-cons ())
            (col-index 0)
            (row-index 0)
            (row-index-str nil))
        (declare (type fixnum col-index row-index)
                 (type (or cons null) row-cons col-cons cell-attributes cell-value-cons)
                 (type (or string null) style type value row-index-str))
        
        (block process
          (prog ((builder (cxml-xmls:make-xmls-builder :include-default-values nil
                                                       :include-namespace-uri nil )))
           l_row
             ;; find the row start
             (klacks:find-element source "row")
                                        ; serialize the whole row (all columns)
             (setf row-cons (klacks:serialize-element source builder))
             ;; pop "row" label and "row" attributes, read "row" attributes in particular the row...
             (pop row-cons)
             (setf row-index-str (the (or string null)
                                      (car (cdr (assoc "r" (pop row-cons) :test #'string= )))))
             (when row-index-str
               (setf row-index (parse-integer row-index-str))
               (when max-row
                 (if (> row-index max-row) (return-from process))))
             ;; skip row if necessary
             (if (< row-index initial-row)
                 (go l_row))
             (funcall row-begin-function row-index)
             ;; iterate over each col-cons, all left is columns
           l_column
             ;; clear values
             (setf style nil
                   value nil
                   type nil
                   col-index 0)
             (incf col-index)
             (setf col-cons (cdr (pop row-cons))) ; column info, like  (("t" "s") ("s" "3") ("r" "A6")) ("v" NIL "18"))
             ;; skip column if not in column-list 
             (unless (and column-list
                          (not (member col-index column-list)))           
               (setf cell-attributes (car col-cons)) ; ("t" "s") ("s" "3") ("r" "A6")
               (setf cell-value-cons (second col-cons)) ; ("v" NIL "18")
               ;; we process the attributes t and s
               (setf style (second (assoc "s" cell-attributes :test #'string=)))
               (setf type  (second (assoc "t" cell-attributes :test #'string=)))
               ;; now we get the value cons
               (when cell-value-cons ; sometimes there will be no value...
                 (setf value (third cell-value-cons)))
               ;; call function for processing all this cell
               (funcall column-process-function row-index col-index value type style))
             ;; if there are rows in row-cons, we repeat for the next column info
             (if (not (null row-cons)) (go l_column))
           l_rowend
             ;; end of row
             (funcall row-end-function row-index)
             ;; go to next row. EOF should terminate all this.
             (go l_row)))
        (funcall final-function)
        ))))

;; Old version that uses cl-xmlspam
;; (defun %%process-sheet (sheet-struct
;;                         &key
;;                           row-begin-function
;;                           column-process-function
;;                           row-end-function
;;                           final-function
;;                           (max-row nil)
;;                           (initial-row 1)
;;                           (column-list nil))  ; get only selected columns (list of indexes)
;;                           ;(initial-stream-position nil) 
                          
;;   "Generalized Process sheet (as struct)"
;;   (declare (type sheet sheet-struct)
;;            (type fixnum initial-row max-row)
;;            (type function row-begin-function
;;                  column-process-function
;;                  row-end-function
;;                  final-function)
;;            (type (or cons null) column-list))
;;   (let ((filename (sheet-file-name sheet-struct)))
;;     (with-open-file (str-utf filename :direction :input
;;                                       ;; SBCL implementation dependent?
;;                                       :external-format '(:utf-8 :replacement #\?)
;;                                       :element-type *element-type*)
;;       ;; if a stream position has been specified, advance (seek) to such position
;;       ; (when initial-stream-position (file-position str-utf initial-stream-position))
;;       ;; do stuff with cl-xmlspam library
;;       (let ((style nil)
;;             (value nil)
;;             (type "")
;;             (col-index 0)
;;             (row-index 0))
;;         (declare (type fixnum col-index row-index)
;;                  (type string type)
;;                  (type (or string null) style value))
;;         (with-xspam-source str-utf
;;           (block process
;;             (element |sheetData|
;;               (group
;;                (one-or-more
;;                 (element :row
;;                   ;; a row -- start row processing
;;                   (attribute :r         ; row number
;;                     (setf row-index (parse-integer (the string _)))
;;                     ;; exit if max-row has been reached
;;                     (when max-row
;;                       (if (> row-index max-row) (return-from process))))
;;                   (when (>= row-index initial-row)
;;                     ;; clear row values
;;                     (setf style nil
;;                           value nil
;;                           type ""
;;                           col-index 0)
;;                     (funcall row-begin-function row-index)
;;                     (group
;;                      (one-or-more       ; zero or more?
;;                       (element :c
;;                         ;;  a column -- start column procesing
;;                         (incf col-index)
;;                         ;; skip this column if index not in the (optional) column list
;;                         (unless (and column-list
;;                                      (not (member col-index column-list)))
;;                           (optional-attribute :s ; style (optional)
;;                             (setf style (the string _)))
;;                           (optional-attribute :t ; type (optional)
;;                             (setf type (the string _)))
;;                           (optional     ; There can be cells with no value !
;;                            (element :v  ; value tag
;;                              (text      ; store the value
;;                               (setf value (the string _)))))
;;                           (funcall column-process-function row-index col-index value type style)
;;                           ))))
;;                     ;; end row
;;                     (funcall row-end-function row-index)
;;                     )))))))
;;         ;; store last stream position, for future usage
;;         (setf (sheet-last-stream-position sheet-struct) (file-position str-utf))
;;         (funcall final-function)
;;         ))))

(defun process-sheet (sheet-struct &key (max-row nil)
                                        (initial-row 1)
                                        (column-list nil) ;; get only selected columns (list of indexes)
                                        (row-function nil)
                                        (silent nil)
                                        (debug-print nil))
                                        ;(initial-stream-position nil)) ;; for future use...
"Process sheet (as struct). Reads rows from the sheet-struct and returns them as cons.
  Important options: 
  * max-row & initial-row control the rows to fetch. First row = 1. 
  * column-list receives a list of column numbers, so only those columns are fetched. First column = 1.
  * row-function receives a lambda where it gets passed, as an argument, the row's column values (as a list).
                 This can be used, for example, for uploading each row to a DB, etc. 
                 If row-function is defined, then <the row values will not be returned by process-sheet.
"
  (declare (type sheet sheet-struct)
           (type fixnum initial-row max-row)
           (type boolean silent debug-print)
           (type (or cons null) column-list)
           (type (or (function (cons)) null) row-function))
  (let* ((date-formats (sheet-date-formats sheet-struct))
         (number-formats (sheet-number-formats sheet-struct))
         (unique-strings (sheet-unique-strings sheet-struct))
         ;; pre-process decisions on how to handle cells
         (decision-vector (%get-col-parse-decisions-vector date-formats number-formats))
         (col-cons nil)
         (row-cons nil))
    (declare (type (or cons null) row-cons col-cons date-formats number-formats)
             (type (vector string) unique-strings)
             (type (vector fixnum) decision-vector))           
    (flet 
        
        ((row-begin-function (row-index)
           (declare (type fixnum row-index))
           (setf col-cons nil)
           (when debug-print (format t "~%Row Number = ~D || " row-index)))
         (column-process-function (row-index col-index value type style)
           (declare (type fixnum col-index)
                    (ignore row-index)
                    (type (or string null) value type style ))
           (when debug-print (format t "|~D " col-index))
           ;; log format ID, type and value
           (when debug-print (format t "(~D,~A,~A) "
                                     ;; number format type ID 
                                     (car (elt number-formats (parse-integer style)))
                                     type
                                     value))
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
                       :initial-row initial-row
                       :column-list column-list))))
                       ;:initial-stream-position initial-stream-position

 


;; simple helper
(defun sheet-first-row (sheet-struct)
  "Obtain first row of sheet"
  (process-sheet sheet-struct :max-row 1 :silent T))


;; helper for column info
(defstruct column-info
  format-id
  format-keyword
  type)

(defun report-cells-type-change (sheet-struct &key (max-row nil)
                                                   (initial-row 1)
                                                   (column-list nil)) ;; get only selected columns (list of indexes)
  "Inspect how cell type changes from row to row. Returns number of changes."
  (declare (type sheet sheet-struct)
           (type fixnum initial-row max-row)
           (type (or cons null) column-list))
  (let* ((number-formats (sheet-number-formats sheet-struct))
         (info nil)
         (number-of-changes 0)
         (blank-row T))
    ;; (col-count 0))
    (declare (type (or cons null) info number-formats)
             (type (boolean) blank-row)    
             (type fixnum number-of-changes))           
    (flet ((row-begin-function (row-index)
             (declare (type fixnum row-index) (ignore row-index))
             (setf blank-row T))        ; reset blank value (row) flag
                                       
           (column-process-function (row-index col-index value type style)
             (declare (type fixnum col-index)
                      (type (or string null) value type style ))
             ;; compare style, type  for this column, versus previous.
        
             ;; detect blank values - set to T if all row is blank or blank strings
             (setf blank-row (and blank-row
                                  (or (null value) (eq 0 (length value)))))
             (let* ((fmt-id (car (elt number-formats (parse-integer style))))
                    (format-keyword (%get-format-type fmt-id))
                    (cinfo (cdr (assoc col-index info)))
                    (equal?
                      (and cinfo ;; there is info for this column
                           (equal (column-info-format-keyword cinfo) format-keyword)
                           (equal (column-info-type cinfo) type))))
               (declare (type fixnum fmt-id)
                        (type (or symbol null) format-keyword)
                        (type (or column-info null) cinfo))
               (unless equal?
                 ;; different style for this column, protest:
                 (incf number-of-changes)
                 (format t "Row ~D: Column ~D changed to format->~A(~D) | type-> ~A ~%"
                         row-index col-index format-keyword fmt-id type)
                 ;; store new style info for this column and remove old.
                 (setf info
                       (cons (cons col-index (make-column-info :type type
                                                               :format-id fmt-id
                                                               :format-keyword format-keyword))
                             (remove (assoc col-index info) info))))))
            
           (row-end-function (row-index)
             (declare (type fixnum row-index))
             (when blank-row
               (format t "~%*********** WARNING : row ~D all blank. ************ ~%~%" row-index)
               ))
           (final-function  ()
             (format t "Number of changes: ~D" number-of-changes)
             number-of-changes
             ))
      ;; call to the function that performs the actual process
      (%%process-sheet sheet-struct
                       :row-begin-function #'row-begin-function
                       :column-process-function #'column-process-function
                       :row-end-function #'row-end-function
                       :final-function #'final-function
                       :max-row max-row
                       :initial-row initial-row
                       :column-list column-list))))


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



;; --test
;;(defparameter *f* "C:\\Users\\fegoavil010\\Documents\\archivos bulk\\GOOD_YEAR\\kardex\\temp.xlsx" )

           
