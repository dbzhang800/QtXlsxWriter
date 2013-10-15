Documentation: http://qtxlsx.debao.me

QtXlsx is a library that can write Excel files. It doesn't require Microsoft Excel and can be used in any platform that Qt5 supported.

## Getting Started

### Usage(1): Use source code directly

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

**Note**: If you like, you can copy all files from *src/xlsx* to your application's source path. Then add following line to your project file:

```
    include(qtxlsx.pri)
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

### Usage(2): Use Xlsx as Qt5's addon module

**Note**: Perl is needed.

* Download the source code.

* Put the source code in any directory you like. At the toplevel directory run

```
    qmake
    make
    make install
```

The library, the header files, and the feature file will be installed to your system.

* Add following line to your qmake's project file:

```
    QT += xlsx
```

* Then, using Qt Xlsx in your code

## References

* https://github.com/jmcnamara/XlsxWriter
* http://officeopenxml.com/anatomyofOOXML-xlsx.php
* http://www.libxl.com
* http://closedxml.codeplex.com/
* http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX-0.71/
* http://www.codeproject.com/Articles/208075/How-to-read-and-write-xlsx-Excel-2007-file-Part-I
* http://www.codeproject.com/Articles/210014/How-to-read-and-write-xlsx-Excel-2007-file-Part-II
* http://epplus.codeplex.com/
* http://excelpackage.codeplex.com/
* http://spreadsheetlight.com/
