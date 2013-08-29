INCLUDEPATH += $$PWD
DEPENDPATH += $$PWD

QT += core gui gui-private
!build_xlsx_lib:DEFINES += XLSX_NO_LIB

HEADERS += $$PWD/xlsxdocprops_p.h \
    $$PWD/xlsxrelationships_p.h \
    $$PWD/xlsxutility_p.h \
    $$PWD/xlsxsharedstrings_p.h \
    $$PWD/xmlstreamwriter_p.h \
    $$PWD/xlsxcontenttypes_p.h \
    $$PWD/xlsxtheme_p.h \
    $$PWD/xlsxformat.h \
    $$PWD/xlsxworkbook.h \
    $$PWD/xlsxstyles_p.h \
    $$PWD/xlsxworksheet.h \
    $$PWD/zipwriter_p.h \
    $$PWD/xlsxpackage_p.h \
    $$PWD/xlsxworkbook_p.h \
    $$PWD/xlsxworksheet_p.h \
    $$PWD/xlsxformat_p.h \
    $$PWD/xlsxglobal.h

SOURCES += $$PWD/xlsxdocprops.cpp \
    $$PWD/xlsxrelationships.cpp \
    $$PWD/xlsxutility.cpp \
    $$PWD/xlsxsharedstrings.cpp \
    $$PWD/xmlstreamwriter.cpp \
    $$PWD/xlsxcontenttypes.cpp \
    $$PWD/xlsxtheme.cpp \
    $$PWD/xlsxformat.cpp \
    $$PWD/xlsxstyles.cpp \
    $$PWD/xlsxworkbook.cpp \
    $$PWD/xlsxworksheet.cpp \
    $$PWD/zipwriter.cpp \
    $$PWD/xlsxpackage.cpp
