(in-package :cl-user)

(defpackage :lisp-xl.metadata
  (:use #:cl)
  (:export #:get-entry
           #:list-sheets
           #:get-unique-strings-vector
           #:get-number-formats
           #:%get-date-formats
           ))

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

