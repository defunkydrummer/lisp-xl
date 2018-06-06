(in-package :lisp-xl)

;(declaim (optimize (speed 3) (debug 0) (safety 0) (space 0)))

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
;; default decision if 'style' not available
(defparameter *decision-default* *decision-parse-string*)

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
          
          for decision =
          (cond
            ((equal format-type :string) *decision-parse-string*)
            ((member format-type *number-formats*) *decision-parse-number*)
            ;; unsupported format: return string.
            ((member format-type *unsupported-formats*) *decision-parse-string*)
            (t *decision-try-lisp-reader*)) ; note! 
          do (setf (aref vector vector-index) decision)
             (incf vector-index))
    vector
    ))

;; speedups for %read-col
(declaim (ftype (function ((or string null ) ;value
                                             (or string null ) ;type
                                             (or string null ) ; style
                                             (vector string) ; unique-strings
                                             (simple-array fixnum)) t) ; decision vector
                %read-col))
;; inlining
(declaim (inline %read-col))
;;(declaim (notinline %read-col))
(defun %read-col (value type style unique-strings decision-vector)
  "Inner function for reading(parsing) a column"
  (declare (type (or string null) value style type)
           (type (vector string) unique-strings)
           (type (simple-array fixnum) decision-vector))
  (let ((decision ; decision of the decision-vector
          (if (null style)
              *decision-default* ;default decision if no style 
              (aref decision-vector (parse-integer style)))))
    (declare (type fixnum decision))
    (handler-case 
        (cond ((null value) nil) ; value is nil (case for blank cells)
              ;; ((string= type "e") (intern value "KEYWORD"))
              ((not (null type))
               (cond 
                 ((string= type "str") value) ;; CALCULATED STRING, we just return the value
                 ;; string from the 'unique-strings' file?
                 ((string= type "s")
                  (if
                   ;; number?
                   (eq decision *decision-parse-number*) (read-from-string value)
                   ;; else - read string from table
                   (aref (the (vector string) unique-strings)
                         (the fixnum (parse-integer value)))))))
              ;; date, etc
              ;; return as string!
              (t value)) 
      (error () "PARSE ERROR") ; return "ERROR" if parse errors were found
      )))
