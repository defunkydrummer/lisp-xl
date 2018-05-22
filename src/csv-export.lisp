(in-package :cl-user)
(defpackage :lisp-xl-csv
  (:use #:cl :lisp-xl :cl-csv)
  (:export :excel-to-csv))
(in-package :lisp-xl-csv)

;; Simple example on how to create a CSV based on the XLSX.

(defun write-row-to-csv (row-data str)
  (declare (type cons row-data)
           (type stream str))
  (cl-csv:write-csv-row
     row-data :stream str ))

(defun excel-to-csv (in-file out-file sheet-index &key (initial-row 1) (max-row nil) (silent T))
  "Convert XLSX file to CSV file."
  (with-open-file (str out-file :direction :output :if-exists :error
                       :if-does-not-exist :create)
    
    (flet ((writer (row-data)
             (declare (type cons row-data))
             (write-row-to-csv row-data str)))
      (with-open-excel-sheet (in-file sheet-index sheet silent)
        (unless silent (format t "Converting to CSV output file [~A] ...~%" out-file))
        (process-sheet sheet
                       :initial-row initial-row
                       :max-row max-row
                       :silent silent
                       :row-function #'writer)))))

