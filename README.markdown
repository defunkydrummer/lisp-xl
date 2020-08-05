# lisp-xl

Common Lisp Library for working with  Microsoft Excel XLSX files, for example for ETL (extract-transform-load) and data cleansing purposes. This library is geared for working with large files (i.e. 100MB excel files).

It does *not* load the whole file into RAM, thus, processing large files is possible. However, be careful with files that contain a lot of "unique strings" (strings that do not repeat frequently on the file), for they need to be loaded onto RAM. 

**NOTE:** This library is still under development, API might change slightly over time. 

## Reason for being

In the business world, when your request a "data dump" from a customer's system,  you usually want the customer to give you the data in the form of compressed, tidy CSV files. In real life, often you will get a huge XLSX files instead, where for example the first row isn't the header row, blank rows here and there, etc. In short, files that aren't ready for uplaoding... So, read the mundane Microsoft XLSX files using the celestial programming language, Common Lisp!

## Features

* Does NOT load the whole file into RAM.
* Loads and uncompresses the XLSX sheet file only once, to save time when reading it more than once.

## TODO

* Parse numbers and dates correctly. Currently, it will not attempt to parse most numbers and dates, but return them as strings.
* Performance improvements.

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

;; delete temp file / close sheet
(close-sheet *s*)

```

Another way is using the macro "With open-excel-sheet", which opens the sheet, performs a block, and then closes the sheet, thus deleting the temporary file. 

```common-lisp

(with-open-excel-sheet (*f* 1 sheet)
           ;; do stuff with 'sheet'
           ;; ...
           )

```

To do useful stuff, you would create a function that will process each row and do something useful with it (Each row will be received as a list of values).

```common-lisp

 (with-open-excel-sheet (*f* 1 sheet nil)
           ;; our "useful" function that will process each row
           (flet ((f (row) 
                    (print row)))
             ;; call #'f at each row, skip first row
             (process-sheet sheet :max-row 10 :initial-row 2 :row-function #'f)))
```

Or you can work on all the sheets in the file. For example:

```common-lisp

(with-all-excel-sheets (file sh index name) ;bind sheet index and name 
    (format t "~% Working on sheet ~A ~%" name)
    ;; do stuff...
    )
    
```

### CSV Export

CSV Export is available by using  `lisp-xl-csv:excel-to-csv`, at csv-export.lisp. 

## Uses

* CXML (Closure-XML)
* ZIP (xlsx files are zipped)
* xmls
* uiop 
* cl-ppcr

## Caveats

The software is still work in progress, it works, but for example there are some issues: 

* See TODO.

The code itself could have been implemented in a much shorter and simple way, however with lower performance. Performance is still one of my main concerns. A current performance bottleneck is the ZIP expansion of the XLSX file. The "deflate" function used is IMO slow and could be reimplemented. 

## Installation

For CL newbies (welcome!): 

Download lisp-xl, copy to "local-projects", then use quicklisp to load the rest of the required libs and compile:

```common-lisp
(ql:quickload :lisp-xl)
```
## Author

Flavio Egoavil (f_egoavil at microsoft's ancient, high temperature mail system.)

## Acknowledgements

* Functions for reading the XLSX file metadata taken from the xlsx library by Carlos Ungil.
* Akihide Nano, for putting Carlos' library in GitHub.

For working with *small* XLSX files in a Cell-oriented way (i.e. "A1", "B4", etc), you might want to take a look at Carlos library:
[https://gitlab.common-lisp.net/cungil/xlsx](https://gitlab.common-lisp.net/cungil/xlsx)

Japanese users may want to use Akihide Nano(a-nano) library which is based on Carlos Ungil's:
https://github.com/a-nano/xlsx
This one also includes returning the XLSX file as a-list or p-list.

## License

Licensed under the MIT license.
