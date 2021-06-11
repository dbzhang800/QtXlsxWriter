Documentation: http://qtxlsx.debao.me

QtXlsx is a library that can read and write Excel files. It doesn't require Microsoft Excel and can be used in any platform that Qt5 supported.
The library can be used to

* Generate a new .xlsx file from scratch
* Extract data from an existing .xlsx file
* Edit an existing .xlsx file
 
## Getting Started

> * For linux user, if your Qt is installed through package manager tools such "apt-get", make sure that you have installed the Qt5 develop package *qtbase5-private-dev*

### Usage(1): Use Xlsx as Qt5's addon module

#### Building the module

> **Note**: Perl is needed in this step.

* Download the source code.

* Put the source code in any directory you like

* Go to top directory of the project in a terminal and run

```
    qmake
    make
    make install
```

The library, the header files, and others will be installed to your system.

> ```make html_docs``` can be used to generate documentations of the library, and ```make check``` can be used to run unit tests of the library.

#### Using the module

* Add following line to your qmake's project file:

```
    QT += xlsx
```

* Then, using Qt Xlsx in your code

```cpp
    #include <QtXlsx>
    int main()
    {
        QXlsx::Document xlsx;
        xlsx.write("A1", "Hello Qt!");
        xlsx.saveAs("Test.xlsx");
        return 0;
    }
```

### Usage(2): Use source code directly

The package contains a **qtxlsx.pri** file that allows you to integrate the component into applications that use qmake for the build step.

* Download the source code.

* Put the source code in any directory you like. For example, 3rdparty:

```
    |-- project.pro
    |-- ....
    |-- 3rdparty\
    |     |-- qtxlsx\
    |     |
```

* Add following line to your qmake project file:

```
    include(3rdparty/qtxlsx/src/xlsx/qtxlsx.pri)
```

> **Note**: If you like, you can copy all files from *src/xlsx* to your application's source path. Then add following line to your project file:

> ```
    include(qtxlsx.pri)
```

> **Note**: If you do not use qmake, you need to define the following macro manually

> ```
    XLSX_NO_LIB
```


* Then, using Qt Xlsx in your code

```cpp
    #include "xlsxdocument.h"
    int main()
    {
        QXlsx::Document xlsx;
        xlsx.write("A1", "Hello Qt!");
        xlsx.saveAs("Test.xlsx");
        return 0;
    }
```

## References

* http://www.ecma-international.org/publications/standards/Ecma-376.htm
* http://www.iso.org/iso/catalogue_detail?csnumber=51463
* http://msdn.microsoft.com/en-us/library/ee908652%28v=office.12%29.aspx
* http://www.datypic.com/sc/ooxml/

### General

* https://github.com/jmcnamara/XlsxWriter
* http://openpyxl.readthedocs.org
* http://officeopenxml.com/anatomyofOOXML-xlsx.php
* http://www.libxl.com
* http://closedxml.codeplex.com/
* http://epplus.codeplex.com/
* http://excelpackage.codeplex.com/
* http://spreadsheetlight.com/

### Number formats

* http://msdn.microsoft.com/en-us/library/ff529356%28v=office.12%29.aspx
* http://www.ozgrid.com/Excel/excel-custom-number-formats.htm
* http://stackoverflow.com/questions/894805/excel-number-format-what-is-409
* http://office.microsoft.com/en-001/excel-help/create-a-custom-number-format-HP010342372.aspx

### Formula

* http://msdn.microsoft.com/en-us/library/ff533995%28v=office.12%29.aspx
* http://msdn.microsoft.com/en-us/library/dd906358%28v=office.12%29.aspx
* http://homepages.ecs.vuw.ac.nz/~elvis/db/Excel.shtml
* http://ewbi.blogs.com/develops/2004/12/excel_formula_p.html
