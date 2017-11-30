# lisp-xl

* Work in progress, but you can try it already... 

Common Lisp Library for working with  Microsoft Excel XLSX files, for example for ETL (extract-transform-load) and data cleansing purposes. This library is geared for working with large files (i.e. 100MB excel files).

It does *not* load the whole file into RAM, thus, processing large files is possible.

## Reason for being

In the business world, when your request a "data dump" from a customer's system,  you usually want the customer to give you the data in the form of compressed, tidy CSV files. In real life, often you will get a huge XLSX files instead, where for example the first row isn't the header row, blank rows here and there, etc. In short, files that aren't ready for uplaoding... So, read the mundane Microsoft XLSX files using the celestial programming language, Common Lisp!

You can use the (report-cells-type-change) function to take a look at the presence of blank lines or unexpected cell format change in the whole XLSX file.  

## Features

* Does NOT load the whole file into RAM.
* Report of "offending" cells; that is, cells that change data type along the way. 
* Report of unexpected blank rows
* Loads and uncompresses the XLSX sheet file only once, to save time.
* Able to select/load only a range of rows, and/or only a selected range of columns.

## TODO

* Parsing of date values correctly
* Free burger for each user

## Usage

Typical process is as follows:

1. Load the sheet, creating a struct that holds info about the sheet. For this you need the file path, and the name/index of the sheet (1 = first sheet).

2. Query the sheet using the provided functions. 

```common-lisp

;; load lib
(ql:quickload :lisp-xl)
(in-package :lisp-xl)

(defparameter *s* (read-sheet *f* 1)) ; read first sheet of file *f*
(process-sheet *s* :max-row 100) ;; obtain some rows as cons

;; Do a report on which cells/rows change format along the file.
(report-cells-type-change *s* :max-row 100)

;; delete temp file / close sheet
(close-sheet *s*)

```

Another way is using the macro "With open-excel-sheet", which opens the sheet, performs a block, and then closes the sheet, thus deleting the temporary file. 

```common-lisp

(with-open-excel-sheet (*f* 1 sheet)
           ;; do stuff with 'sheet'
           (report-cells-type-change
             sheet :max-row 20))

```

To do useful stuff, you would create a function that will process each row and do something useful with it (Each row will be received as a list of values).

```common-lisp

 (with-open-excel-sheet (*f* 1 sheet nil)
           ;; our "useful" function that will process each row
           (flet ((f (row) 
                    (print row)))
             ;; call #'f at each row
             (process-sheet sheet :max-row 10 :row-function #'f)))
```

## Uses

* cl-xmlspam (download a copy from my GitHub)
* ZIP (xlsx files are zipped)
* xmls
* uiop (

## Installation

For CL newbies (welcome!): 
Make sure you have the cl-xmlspam library downloaded; you can copy it to quicklisp's "local-projects" directory. 
Download lisp-xl, copy to "local-projects", then use quicklisp to load the rest of the required libs and compile:

```common-lisp
(ql:quickload :lisp-xl)
```

## Acknowledgements

* Functions for reading the XLSX file metadata taken from the xlsx library by Carlos Ungil.
* Akihide Nano, for putting Carlos' library in GitHub.

For working with *small* XLSX files in a Cell-oriented way (i.e. "A1", "B4", etc), you might want to take a look at Carlos library:
https://gitlab.common-lisp.net/cungil/xlsx](https://gitlab.common-lisp.net/cungil/xlsx)

Japanese users may want to use Akihide Nano(a-nano) library which is based on Carlos Ungil's:
https://github.com/a-nano/xlsx
This one also includes returning the XLSX file as a-list or p-list.

## License

Licensed under the MIT license.
