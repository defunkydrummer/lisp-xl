# lisp-xl

* Work in progress! *

Common Lisp Library for working with  Microsoft Excel XLSX files, hopefully for ETL (extract-transform-load) and data cleansing purposes. This library is geared for working with large files (i.e. 100MB excel files).

## Reason for being

In the business world, you usually want customers to give you data in the form of compressed, tidy CSV files, but they will send you huge XLSX files instead, files that aren't ready for uplaoding... So, read the mundane Microsoft XLSX files using the celestial programming language, Common Lisp!

## Usage

Typical process is as follows:

1. Load the sheet into a sheet struct. For this you need the file path, and the name/index of the sheet (1 = first sheet).

2. Query the sheet using the provided functions. 

```common-lisp

(ql:quickload :lisp-xl)
(defparameter *s* (read-sheet *f* 1)) ; read first sheet of file *f*
(process-sheet *s* :max-row 100) ;; obtain rows

```
## Uses

* ZIP (xlsx files are zipped)
* babel (for utf8 decoding)
* babel-stream
* xmls
* cl-xmlspam (download a copy from my GitHub)

## Installation

Download, copy to quicklisp's "local-projects" dir, then use quicklisp to load and compile:

```common-lisp
(ql:quickload :lisp-xl)
```

## Acknowledgements

* Some core xlsx metadata read functions taken from the xlsx library by Carlos Ungil
* Akihide Nano, for putting said library in GitHub

For working with *small* XLSX files in a Cell-oriented way (i.e. "A1", "B4", etc), you might want to take a look at Carlos library:
https://gitlab.common-lisp.net/cungil/xlsx](https://gitlab.common-lisp.net/cungil/xlsx)

Japanese users may want to use Akihide Nano(a-nano) library which is based on Carlos Ungil's:
https://github.com/a-nano/xlsx
This one also includes returning the XLSX file as a-list or p-list.

## License

Licensed under the MIT license.
