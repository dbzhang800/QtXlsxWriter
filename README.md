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

* Open the qtxlsx.pro file using Qt Creator
 * Build the project.
 * Play with the examples provided by Qt Xlsx if you like.

* Go to the build directory of the project in a terminal and run

```
   make install
```

The library, the header files, and others will be installed to your system.

> **Note**: If you don't want to use Qt Creator, you can run following command
 at the toplevel directory of the project

> ```
    qmake
    make
    make install
```

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

### General

* https://github.com/jmcnamara/XlsxWriter
* http://openpyxl.readthedocs.org
* http://officeopenxml.com/anatomyofOOXML-xlsx.php
* http://www.libxl.com
* http://closedxml.codeplex.com/
* http://epplus.codeplex.com/
* http://excelpackage.codeplex.com/
* http://spreadsheetlight.com/
* http://www.schemacentral.com/sc/ooxml/

### Number formats

* http://msdn.microsoft.com/en-us/library/ff529356%28v=office.12%29.aspx
* http://www.ozgrid.com/Excel/excel-custom-number-formats.htm
* http://stackoverflow.com/questions/894805/excel-number-format-what-is-409
* http://office.microsoft.com/en-001/excel-help/create-a-custom-number-format-HP010342372.aspx
