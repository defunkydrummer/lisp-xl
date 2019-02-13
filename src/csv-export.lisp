(in-package :cl-user)
(defpackage :lisp-xl-csv
  (:use #:cl :lisp-xl :cl-csv
        :lisp-xl.util)
  (:export :excel-to-csv
   :files-to-csv ))
(in-package :lisp-xl-csv)

;; Simple example on how to create a CSV based on the XLSX.

(defun write-row-to-csv (row-data str)
  (declare (type cons row-data)
           (type stream str))
  (cl-csv:write-csv-row
     row-data :stream str ))


(defun rename-to-csv (file)
  "Obtain renamed file appending .csv extension to original file"
  (format nil "~A.csv" (uiop:physicalize-pathname file)))

(defun excel-to-csv (in-file out-file sheet-index &key (initial-row 1) (max-row nil) (silent T)
                                                       (output-format :utf-8))
  "Convert XLSX file to CSV file. 
Null for out file: rename in-file
Null for sheet-index: use first sheet.
Note: Strings with CR or LF will have them replaced by spaces."
  (when (null out-file)
    (setf out-file (rename-to-csv in-file)))
  (with-open-file (str out-file :direction :output :if-exists :error
                                :if-does-not-exist :create
                                :external-format output-format)
    
    (flet ((writer (row-data)
             (declare (type cons row-data))
             (write-row-to-csv (nfilter-eol row-data) str)))
      (with-open-excel-sheet (in-file sheet-index sheet silent)
        (unless silent (format t "Converting to CSV output file [~A] ...~%" out-file))
        (process-sheet sheet
                       :initial-row initial-row
                       :max-row max-row
                       :silent silent
                       :row-function #'writer)))))

(defun files-to-csv (files &key output-format
                                remove-header-butfirst)
  "Convert all files(assumed excel files) to CSV.
remove-header-butfirst: strip the header in all files except the first.
Processes only the first sheet of each file."
  (when remove-header-butfirst
    (format t "Note: First file will keep its header: ~A~%" (first files)))
  (loop for f in files
        for first = (equal f (first files))
        do
        (format t "Processing file: ~A~%" f)
        (excel-to-csv f nil 1
                      :initial-row
                      (if (and remove-header-butfirst
                               (not first))
                          2 ;;remove header
                          1)
                      :output-format output-format)))


