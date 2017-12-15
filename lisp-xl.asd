(in-package :cl-user)
(defpackage lisp-xl-asd
  (:use :cl :asdf))
(in-package :lisp-xl-asd)

(defsystem lisp-xl
  :version "0.1"
  :author "Flavio Egoavil"
  :license "MIT license"
  :depends-on (:zip
               :cxml 
               :uiop
               :xmls)
  :components ((:module "src"
                :components
                ((:file "package")
                 (:file "metadata")
                 (:file "col-process")
                 (:file "main"))))
  :description "Read Microsoft XLSX files using Common Lisp"
  :long-description
  #.(with-open-file (stream (merge-pathnames
                             #p"README.markdown"
                             (or *load-pathname* *compile-file-pathname*))
                            :if-does-not-exist nil
                            :direction :input)
      (when stream
        (let ((seq (make-array (file-length stream)
                               :element-type 'character
                               :fill-pointer t)))
          (setf (fill-pointer seq) (read-sequence seq stream))
          seq)))
  :in-order-to ((test-op (test-op xlsx-test))))
