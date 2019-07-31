(in-package :lisp-xl.metadata)

;;(declaim (optimize (speed 1) (debug 3) (safety 3)))

;; number of entries for unique-strings to be allocated initially.-
;; used each time we do vector-push-extend on the unique strings table.
(defparameter *initial-unique-strings-array-size* 1000) 

;; From Carlos Ungil
(defun get-entry (name zip)
  "Internal use, gets a entry inside the ZIP file. XLSX files are zip files."
  (let ((entry (zip:get-zipfile-entry name zip)))
    ;; get-zipfile-entry: Return an entry handle for the file called name.
    (when entry (xmls:parse (babel:octets-to-string (zip:zipfile-entry-contents entry))))))

;; From Carlos Ungil
(defun get-relationships(zip)
  "Use XLSX relationship file."
  (loop for rel in (xmls:xmlrep-find-child-tags 
		    :relationship (get-entry "xl/_rels/workbook.xml.rels" zip))
        collect (cons (xmls:xmlrep-attrib-value "Id" rel)
                      (xmls:xmlrep-attrib-value "Target" rel))))

;; From Carlos Ungil
;; defunkydrummer-> converted to return an adjustable vector (!) 
(defun get-unique-strings-vector (zip)
  "Retrieves VECTOR  of unique strings used on the file. Will be used later for retrieving string values from the xlsx file."
  (let ((vector (make-array *initial-unique-strings-array-size* :element-type 'string
                                                                :fill-pointer 0
                                                                :adjustable t)))
    (loop for str in (xmls:xmlrep-find-child-tags :si (get-entry "xl/sharedStrings.xml" zip))
          for x = (xmls:xmlrep-find-child-tag :t str)
          do
          (cond ((equal (xmls:node-attrs x) '(("space" "preserve")))
                 (vector-push-extend " "
                                     vector *initial-unique-strings-array-size*))
                ((xmls:xmlrep-children x)
                 (vector-push-extend (xmls:xmlrep-string-child x)
                                     vector *initial-unique-strings-array-size*))
                (t (format t "Warning: Strange entry on unique strings? ~A" x)
                   (vector-push-extend " "
                                       vector *initial-unique-strings-array-size*))))
    vector))

;; From Carlos Ungil
;; Rewriten 2018-06-04, will need further review again. 
(defun get-number-formats (zip)
  (declare (optimize (safety 3)))
  "Retrieves number formats from the 'styles' xml file, which contains style info for the
  excel file."
  (prog ()
     (let* ((style-file (get-entry "xl/styles.xml" zip))
            (numfmts (and style-file
                          (xmls:xmlrep-find-child-tag
                           :numFmts
                           style-file nil)))
            (numfmt-list (and numfmts (xmls:xmlrep-find-child-tags
                                       :numFmt
                                       numfmts))))
       (unless numfmts
         ;; no number format information
         (format t "Lisp-xl warning: File has no number format information. Proceed with caution! ~%")
         ;; return the empty list
         nil)
       (let*
           ((format-codes
              (loop for fmt in numfmt-list
                    for fmt-id-str =
                    ;; handle xmls where no numFmtId info is there
                    (handler-case
                        (xmls:xmlrep-attrib-value "numFmtId" fmt nil)
                      (error nil))
                    when fmt-id-str 
                    collect
                    (cons (parse-integer fmt-id-str)
                          (xmls:xmlrep-attrib-value "formatCode" fmt))))
            (styles-list (and style-file
                              (xmls:xmlrep-find-child-tags
                               :xf (xmls:xmlrep-find-child-tag
                                    :cellXfs style-file)))))
         (when styles-list 
           (loop for style in styles-list
                 for raw-id = (handler-case
                                  (xmls:xmlrep-attrib-value "numFmtId" style nil)
                                (error nil))
                 when raw-id
                 collect
                 (let ((fmt-id (parse-integer raw-id)))
                   (cons fmt-id
                         (if (< fmt-id 164) :built-in
                             (cdr (assoc fmt-id format-codes)))))))))))

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
;; (defun %excel-date (int)
;;   "Decode the excel date values into something we can parse."
;;   (declare (type fixnum int))
;;   (the string 
;;        (apply #'format nil "~D-~2,'0D-~2,'0D"
;;               (reverse (subseq (multiple-value-list (decode-universal-time (* 24 60 60 (- int 2)))) 3 6)))))


;; From Carlos Ungil
(defun list-sheets (file)
  "Retrieves the id and name of the worksheets in the .xlsx/.xlsm file."
  (zip:with-zipfile (zip file)
    (loop for sheet in (xmls:xmlrep-find-child-tags
			:sheet (xmls:xmlrep-find-child-tag
				:sheets (get-entry "xl/workbook.xml" zip)))
          with rels = (get-relationships zip)
          for sheet-id = (xmls:xmlrep-attrib-value "sheetId" sheet)
          for sheet-name = (xmls:xmlrep-attrib-value "name" sheet)
          for sheet-rel = (xmls:xmlrep-attrib-value "id" sheet)
          collect (list (parse-integer sheet-id)
                        sheet-name 
                        (cdr (assoc sheet-rel rels :test #'string=))))))

(defun list-sheet-indexes (file)
  "List all valid indexes for the sheet, in order of appereance"
  (mapcar #'car (list-sheets file)))

(defun sheet-name (file index)
  "Obtain the name (string) of the specific sheet."
  (second (assoc index (list-sheets file))))
