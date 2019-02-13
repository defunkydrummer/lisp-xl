(in-package :lisp-xl.util)

(eval-when (:compile-toplevel :load-toplevel :execute)
  (defparameter *regex* (cl-ppcre:create-scanner "^[A-Z]+"
                                                 :case-insensitive-mode nil
                                                 :single-line-mode T)
    "Regex to capture column number (alphanumerical, i.e. 'AA')"))

(defun get-column-label (str)
  "get column label ('AA33' -> 'AA' etc)"
  (declare (type string str))
  (the (or null string)
       (cl-ppcre:scan-to-strings *regex* str)))

(declaim (inline char-substract))
(defun char-substract (c)
  "Substract character from A and return integer. A=1, B=2, etc."
  (declare (type character c))
  (the fixnum
       (1+ (- (char-code c) (char-code #\A)))))

(declaim (inline get-column-number))
(defun get-column-number (label)
  "Convert column label ('AA') to number starting from 1 for A.
NOTE: Requires label to be UPPERCASE.
(This is converting excel column encoding to integer.)
A -> 1
B -> 2
Z -> 26
AA -> 27
AB -> 28
BA -> 53
etc"
  (declare (type string label))
  (setf label (reverse label))
  (the fixnum
       (loop 
         for index of-type fixnum from 0 to (1- (length label))
         for weight of-type fixnum = (expt 26 index)
         summing
         (* weight (char-substract (elt label index))))))

(defun excel-position-to-column-number (str &aux label) 
  "Convert excel position (i.e. 'AB2') to fixnum that denotes the column index (row is discarded.)"
  (declare (type string str)
           (type (or string null) label))
  (setf label (get-column-label str))
  (unless label
    (error "Invalid label for identifying an excel cell."))
  (the fixnum (get-column-number label)))

(defun label-difference (l1 l2)
  "Difference in number of columns between two excel labels: i.e. (B A) -> 1"
  (declare (type string l1 l2))
  (the fixnum
       (- (excel-position-to-column-number l1)
          (excel-position-to-column-number l2))))
  
(defun %nfilter-eol (str)
  "Replace CR and LF with spaces."
  (declare (type string str))
  (nsubstitute #\Space #\Return str)
  (nsubstitute #\Space #\Newline str))

(defun nfilter-eol (list)
  "Replace CR and LF with spaces in all strings of the list."
  (mapcar (lambda (x)
            (if (stringp x)
                (%nfilter-eol x)
                x))
          list))
