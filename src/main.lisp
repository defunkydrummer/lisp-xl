(in-package :cl-user)
(defpackage :lisp-xl
  (:use #:cl :xspam :lisp-xl.metadata)
  (:export #:list-sheets
           #:read-sheet
           #:process-sheet
           #:reports-cells-type-change
   ))
(in-package :lisp-xl)

;; --- NOTE: debug info goes down with this...
(declaim (optimize (speed 3) (debug 0) (safety 0)))

;; From Carlos Ungil
(defun %get-date-formats (number-formats)
  "Filter (obtain) which formats are Date formats, from the number-formats list"
  (loop for f in number-formats
        for id = (car f)
        for fmt = (cdr f)
        for is-date = 
        (or (<= 14 id 17) ;; built-in: m/d/yyyy d-mmm-yy d-mmm mmm-yy
            (and (stringp fmt) (not (search "h" fmt)) (not (search "s" fmt))
                 (search "d" fmt) (search "m" fmt) (search "y" fmt)))
        collect is-date))

;; From Carlos Ungil
(defun %excel-date (int)
  "Decode the excel date values into something we can parse."
  (apply #'format nil "~D-~2,'0D-~2,'0D"
	 (reverse (subseq (multiple-value-list (decode-universal-time (* 24 60 60 (- int 2)))) 3 6))))

(defun %format-numeric-p (id)
  "Check if Format ID is of a numeric type according to the XLSX standard."
  (declare (type fixnum id))
  (or (<= 1 id 10)
      ;; TODO: Format #11 and #48 are scientific (exponential) notation, what to do with them?
      (<= 37 id 40)))

(defun %read-col (value type style unique-strings date-formats number-formats)
  "Inner function for reading(parsing) a column"
  (declare (type (or string null) value type style)
           (type (or cons null) unique-strings date-formats number-formats))
  (let* ((nstyle (parse-integer style))
         (date? (and (elt date-formats nstyle)))  ; T if style is part of a date format, otherwise NIL
         (number? (%format-numeric-p (car (elt number-formats nstyle))))) ; T if format is numeric
    (declare (type fixnum nstyle)
             (type boolean date? number?))
    (handler-case 
        (cond ((equal type "e") (intern value "KEYWORD")) ;;ERROR
              ((equal type "str") value) ;; CALCULATED STRING
              ;; string from the 'unique-strings' file?
              ((equal type "s")
               (if number?
                   (read-from-string value) ;; read number using Lisp reader
                   ;; else get from unique-strings
                   (nth (parse-integer value) unique-strings)))
              (date? (%excel-date (parse-integer value)))
              ;; Any other case: let Lisp parse the value. 
              (t (read-from-string value)))
      (error () "PARSE ERROR")   ; return "ERROR" if parse errors were found
      )))

(defparameter *element-type* '(unsigned-byte 8))

;; Struct for containing sheet information
(defstruct sheet
  unique-strings
  number-formats
  date-formats 
  data-array ;bulk XML data as array
)  

(defun read-sheet (file sheet &key (silent nil))
  "Read excel file sheet into new sheet struct"
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
    (unless silent (format t "Reading file ~A...~%" file))
    (zip:with-zipfile (zip file)
      ;; get metadata needed for further process
      (setf (sheet-unique-strings sheet-struct) (get-unique-strings zip))
      (setf (sheet-number-formats sheet-struct) (get-number-formats zip))
      (setf (sheet-date-formats sheet-struct)
            (%get-date-formats (sheet-number-formats sheet-struct)))
      ;; read /expand the sheet data itself
      (let ((entry (zip:get-zipfile-entry (format nil "xl/~A" entry-name)
                                          zip)))
        ;; get-zipfile-entry: Return an entry handle for the file called name.
        (when entry
          ;; Load ZIP into buffer array
          (unless silent (format t  "Uncompressing...~%"))
          (setf (sheet-data-array sheet-struct) (zip:zipfile-entry-contents entry))
          (unless silent (format t "Expanded size = ~D bytes ~%" (length (sheet-data-array sheet-struct))))
          ;; return the sheet struct
          sheet-struct
          )))))


(defun %%process-sheet (sheet-struct
                        &key
                          row-begin-function
                          column-process-function
                          row-end-function
                          final-function
                          (max-row nil)
                          (initial-row 1)
                          (column-list nil)) ;; get only selected columns (list of indexes)
  "Generalized Process sheet (as struct)"
  ;; use babel to create an in-memory stream that will convert to utf-8
  (declare (type sheet sheet-struct)
           (type fixnum initial-row max-row)
           (type function row-begin-function
                 column-process-function
                 row-end-function
                 final-function)
           (type (or cons null) column-list))
  (let ((arr (sheet-data-array sheet-struct)))
    (babel-streams:with-input-from-sequence (str-utf arr :element-type *element-type*)
      ;; do stuff with cl-xmlspam library
      (let ((style nil)
            (value nil)
            (type nil)
            (col-index 0)
            (row-index 0))
        (declare (type fixnum col-index row-index)
                 (type (or string null) style type value))
        (with-xspam-source str-utf
          (block process
            (element |sheetData|
              (group
               (one-or-more
                (element :row
                  ;; a row -- start row processing
                  (attribute :r         ; row number
                    (setf row-index (parse-integer (the string _)))
                    ;; exit if max-row has been reached
                    (when max-row
                      (if (> row-index max-row) (return-from process))))
                  (when (>= row-index initial-row)
                    ;; clear row values
                    (setf style nil
                          value nil
                          type nil
                          col-index 0)
                    (funcall row-begin-function row-index)
                    (group
                     (one-or-more       ; zero or more?
                      (element :c
                        ;;  a column -- start column procesing
                        (incf col-index)
                        ;; skip this column if index not in the (optional) column list
                        (unless (and column-list
                                     (not (member col-index column-list)))
                          (optional-attribute :s ; style (optional)
                            (setf style (the string _)))
                          (optional-attribute :t ; type (optional)
                            (setf type (the string _)))
                          (one-of
                           (element :v  ; value tag
                             (text      ; store the value
                              (setf value (the string _)))))
                          (funcall column-process-function row-index col-index value type style)
                          ))))
                    ;; end row
                    (funcall row-end-function row-index)
                    )))))))
        (funcall final-function)
        ))))

(defun process-sheet (sheet-struct &key (max-row nil)
                                        (initial-row 1)
                                        (column-list nil) ;; get only selected columns (list of indexes)
                                        (silent nil)
                                        (debug-print nil))
  "Process sheet (as struct)"
  (declare (type sheet sheet-struct)
           (type fixnum initial-row max-row)
           (type boolean silent debug-print)
           (type (or cons null) column-list))
  (let* ((date-formats (sheet-date-formats sheet-struct))
        (number-formats (sheet-number-formats sheet-struct))
        (unique-strings (sheet-unique-strings sheet-struct))
        (col-cons nil)
        (row-cons nil)
         (row-begin-function (lambda (row-index)
                              (declare (type fixnum row-index))
                              (when debug-print (format t "~%Row Number = ~D || " row-index))))
         (column-process-function (lambda (row-index col-index value type style)
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
                                                    date-formats
                                                    number-formats)
                                         col-cons)))
         (row-end-function (lambda (row-index)
                             (declare (type fixnum row-index) (ignore row-index))
                            (push (nreverse col-cons) row-cons)))
        (final-function (lambda ()
                          (unless silent (print "Finalizing rows..."))
                          (nreverse row-cons)))) ; return rows
    (declare (type (or cons null) row-cons col-cons date-formats number-formats unique-strings))
    ;; call to the function that performs the actual process
    (%%process-sheet sheet-struct
                     :row-begin-function row-begin-function
                     :column-process-function column-process-function
                     :row-end-function row-end-function
                     :final-function final-function
                     :max-row max-row
                     :initial-row initial-row
                     :column-list column-list)))

 


;; simple helper
(defun sheet-first-row (sheet-struct)
  "Obtain first row of sheet"
  (process-sheet sheet-struct :max-row 1 :silent T))


;; helper for column info
(defstruct column-info
  style
  type)

(defun report-cells-type-change (sheet-struct &key (max-row nil)
                                                   (initial-row 1)
                                                   (column-list nil)) ;; get only selected columns (list of indexes)
  "Inspect how cell type changes from row to row. Returns number of changes."
  (declare (type sheet sheet-struct)
           (type fixnum initial-row max-row)
           (type (or cons null) column-list))
  (let* (; (date-formats (sheet-date-formats sheet-struct))
         (number-formats (sheet-number-formats sheet-struct))
         ;(unique-strings (sheet-unique-strings sheet-struct))
         ;; (col-cons nil)
         ;; (row-cons nil)
         (info nil)
         (number-of-changes 0)         
         (row-begin-function (lambda (row-index)
                               (declare (type fixnum row-index) (ignore row-index))
                               nil ))
         (column-process-function (lambda (row-index col-index value type style)
                                    (declare (type fixnum col-index)
                                             (ignore value)
                                             (type (or string null) value type style ))
                                    ;; compare style, type  for this column, versus previous. 
                                    (let* ((cinfo (cdr (assoc col-index info)))
                                           (equal?
                                             (and cinfo ;; there is info for this column
                                                  (equal (column-info-style cinfo) style)
                                                  (equal (column-info-type cinfo) type))))
                                      (unless equal?
                                        ;; different style for this column, protest:
                                        (incf number-of-changes)
                                        (format t "Row ~D: Column ~D changed to format ID->~A | type-> ~A ~%"
                                                row-index col-index
                                                (car (elt number-formats (parse-integer style)))
                                                type)
                                        ;; store new style info for this column and remove old.
                                        (setf info
                                              (cons (cons col-index (make-column-info :style style :type type))
                                                    (remove (assoc col-index info) info)))))                                    
                                    ))
         (row-end-function (lambda (row-index)
                             (declare (type fixnum row-index) (ignore row-index))
                             nil ))
         (final-function (lambda ()
                           (format t "Number of changes: ~D" number-of-changes)
                           number-of-changes
                           ))) 
    (declare (type (or cons null) info number-formats)
             (type fixnum number-of-changes ))
    ;; call to the function that performs the actual process
    (%%process-sheet sheet-struct
                     :row-begin-function row-begin-function
                     :column-process-function column-process-function
                     :row-end-function row-end-function
                     :final-function final-function
                     :max-row max-row
                     :initial-row initial-row
                     :column-list column-list)))


;; -- to be deleted-
;; (defun old-process-sheet (sheet-struct &key (max-row nil)
;;                                         (initial-row 1)
;;                                         (column-list nil) ;; get only selected columns (list of indexes)
;;                                         (silent nil)
;;                                         (debug-print nil))
;;   "Process sheet (as struct)"
;;   ;; use babel to create an in-memory stream that will convert to utf-8
;;   (declare (type sheet sheet-struct)
;;            (type fixnum initial-row max-row)
;;            (type boolean silent debug-print)
;;            (type (or cons null) column-list))
;;   (let ((arr (sheet-data-array sheet-struct))
;;         (date-formats (sheet-date-formats sheet-struct))
;;         (number-formats (sheet-number-formats sheet-struct))
;;         (unique-strings (sheet-unique-strings sheet-struct)))
;;     (babel-streams:with-input-from-sequence (str-utf arr :element-type *element-type*)
;;       ;; do stuff with cl-xmlspam library
;;       (unless silent (print "Processing..."))
;;       (let ((style "")
;;             (value "")
;;             (type "")
;;             (row-cons nil)
;;             (col-cons nil)
;;             (col-index 0)
;;             (row-index 0))
;;         (declare (type fixnum col-index row-index)
;;                  (type (or string null) style type value)
;;                  (type (or cons null) row-cons col-cons))
;;         (with-xspam-source str-utf
;;           (block process
;;             (element |sheetData|
;;               (group
;;                (one-or-more
;;                 (element :row
;;                   ;; a row -- start row processing
;;                   (attribute :r         ; row number
;;                     (setf row-index (parse-integer _))
;;                     ;; exit if max-row has been reached
;;                     (when max-row
;;                       (if (> row-index max-row) (return-from process)))
;;                     (when debug-print (format t "~%Row Number = ~D || " row-index)))
;;                   (when  debug-print (format t "Previous row values: ~A, ~A, ~A~%" style value type ))
;;                   (when (>= row-index initial-row)
;;                     ;; clear row values
;;                     (setf style nil
;;                           value nil
;;                           type nil
;;                           col-index 0
;;                           col-cons nil)
;;                     (group
;;                      (one-or-more       ; zero or more?
;;                       (element :c
;;                         ;;  a column -- start column procesing
;;                         (incf col-index)
;;                         ;; skip this column if index not in the (optional) column list
;;                         (unless (and column-list
;;                                      (not (member col-index column-list)))
;;                           (when debug-print (format t "|~D " col-index))
;;                           (optional-attribute :s ; style (optional)
;;                             (setf style _))
;;                           (optional-attribute :t ; type (optional)
;;                             (setf type _))
;;                           (one-of
;;                            (element :v  ; value tag
;;                              (text      ; store the value
;;                               (setf value _))))
;;                           ;; log format ID, type and value
;;                           (when debug-print (format t "(~D,~A,~A) "
;;                                                     ;; number format type ID 
;;                                                     (car (elt number-formats (parse-integer style)))
;;                                                     type
;;                                                     value))
;;                           ;; store column value
;;                           (push (%read-col value
;;                                            type
;;                                            style
;;                                            unique-strings
;;                                            date-formats
;;                                            number-formats)
;;                                 col-cons)))))
;;                     ;; end row, push it. 
;;                     (push (nreverse col-cons) row-cons)
;;                     )))))))
;;         (unless silent (print "Finalizing rows..."))
;;         (nreverse row-cons)                        ; return rows
;;         ))))


;; to be deleted
;; (defun old-report-cells-type-change (sheet-struct &key (max-row nil)
;;                                         (initial-row 1)
;;                                         (column-list nil) ;; get only selected columns (list of indexes)
                                        
;;                                         )
;;   "Inspect how cell type changes from row to row. Returns number of changes."
;;   ;; use babel to create an in-memory stream that will convert to utf-8
;;   (declare (type sheet sheet-struct)
;;            (type fixnum initial-row max-row)
;;            (type (or cons null) column-list))
;;   (let ((arr (sheet-data-array sheet-struct))
;;         (date-formats (sheet-date-formats sheet-struct))
;;         (number-formats (sheet-number-formats sheet-struct))
;;         (unique-strings (sheet-unique-strings sheet-struct)))
;;     (babel-streams:with-input-from-sequence (str-utf arr :element-type *element-type*)
;;       ;; do stuff with cl-xmlspam library
;;       (let ((style nil)
;;             (value nil)
;;             (type nil)
;;             (info nil)
;;             (number-of-changes 0)
;;             (col-index 0)
;;             (row-index 0))
;;         (declare (type fixnum col-index row-index number-of-changes)
;;                  (type (or string null) style type value)
;;                  (type (or cons null) info ))
                 
;;         (with-xspam-source str-utf
;;           (block process
;;             (element |sheetData|
;;               (group
;;                (one-or-more
;;                 (element :row
;;                   ;; a row -- start row processing
;;                   (attribute :r         ; row number
;;                     (setf row-index (parse-integer _))
;;                     ;; exit if max-row has been reached
;;                     (when max-row
;;                       (if (> row-index max-row) (return-from process))))
;;                   (when (>= row-index initial-row)
;;                     ;; clear row values
;;                     (setf col-index 0)
;;                     (group
;;                      (one-or-more       ; zero or more?
;;                       (element :c
;;                         ;;  a column -- start column procesing
;;                         (incf col-index)
;;                         ;; skip this column if index not in the (optional) column list
;;                         (unless (and column-list
;;                                      (not (member col-index column-list)))
;;                           (optional-attribute :s ; style (optional)
;;                             (setf style _))
;;                           (optional-attribute :t ; type (optional)
;;                             (setf type _))
;;                           (one-of
;;                            (element :v  ; value tag
;;                              (text      ; store the value
;;                               (setf value _))))
;;                           ;; compare style, type  for this column, versus previous. 
;;                           (let* ((cinfo (cdr (assoc col-index info)))
;;                                  (equal?
;;                                    (and cinfo ;; there is info for this column
;;                                         (equal (column-info-style cinfo) style)
;;                                         (equal (column-info-type cinfo) type))))
;;                             (unless equal?
;;                               ;; different style for this column, protest:
;;                               (incf number-of-changes)
;;                               (format t "Row ~D: Column ~D changed to format ID->~A | type-> ~A ~%"
;;                                       row-index col-index
;;                                       (car (elt number-formats (parse-integer style)))
;;                                       type)
;;                               ;; store new style info for this column and remove old.
;;                               (setf info
;;                                     (cons (cons col-index (make-column-info :style style :type type))
;;                                           (remove (assoc col-index info) info)))))
;;                           ;; store column value
;;                           ))))
;;                     )))))))
;;                                         ; return
;;         number-of-changes
;;         ))))

;; --test
(defparameter *f* "C:\\Users\\fegoavil010\\Documents\\archivos bulk\\GOOD_YEAR\\kardex\\temp.xlsx" )

