INCLUDEPATH += $$PWD
DEPENDPATH += $$PWD

QT += core gui gui-private
!build_xlsx_lib:DEFINES += XLSX_NO_LIB

HEADERS += $$PWD/xlsxdocpropscore_p.h \
    $$PWD/xlsxdocpropsapp_p.h \
    $$PWD/xlsxrelationships_p.h \
    $$PWD/xlsxutility_p.h \
    $$PWD/xlsxsharedstrings_p.h \
    $$PWD/xlsxxmlwriter_p.h \
    $$PWD/xlsxcontenttypes_p.h \
    $$PWD/xlsxtheme_p.h \
    $$PWD/xlsxformat.h \
    $$PWD/xlsxworkbook.h \
    $$PWD/xlsxstyles_p.h \
    $$PWD/xlsxworksheet.h \
    $$PWD/xlsxzipwriter_p.h \
    $$PWD/xlsxpackage_p.h \
    $$PWD/xlsxworkbook_p.h \
    $$PWD/xlsxworksheet_p.h \
    $$PWD/xlsxformat_p.h \
    $$PWD/xlsxglobal.h \
    $$PWD/xlsxdrawing_p.h

SOURCES += $$PWD/xlsxdocpropscore.cpp \
    $$PWD/xlsxdocpropsapp.cpp \
    $$PWD/xlsxrelationships.cpp \
    $$PWD/xlsxutility.cpp \
    $$PWD/xlsxsharedstrings.cpp \
    $$PWD/xlsxxmlwriter.cpp \
    $$PWD/xlsxcontenttypes.cpp \
    $$PWD/xlsxtheme.cpp \
    $$PWD/xlsxformat.cpp \
    $$PWD/xlsxstyles.cpp \
    $$PWD/xlsxworkbook.cpp \
    $$PWD/xlsxworksheet.cpp \
    $$PWD/xlsxzipwriter.cpp \
    $$PWD/xlsxpackage.cpp \
    $$PWD/xlsxdrawing.cpp
