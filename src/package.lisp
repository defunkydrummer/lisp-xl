(in-package :cl-user)

(defpackage :lisp-xl.util
  (:use #:cl)
  (:export :excel-position-to-column-number
   :label-difference
   :nfilter-eol))

(defpackage :lisp-xl.metadata
  (:use #:cl)
  (:export #:get-entry
           #:list-sheets
           #:list-sheet-indexes
           #:sheet-name
           #:get-unique-strings-vector
           #:get-number-formats
           #:%get-date-formats
           ))

(defpackage :lisp-xl
  (:use #:cl #|:xspam|# :lisp-xl.metadata
        :lisp-xl.util)
  (:export #:list-sheets
           #:read-sheet
           #:process-sheet
           #:close-sheet
           #:sheet-first-row
           #:with-open-excel-sheet
           #:with-all-excel-sheets
           #:reports-cells-type-change
           ))

