INCLUDEPATH += $$PWD
DEPENDPATH += $$PWD

QT += core gui gui-private
!build_xlsx_lib:DEFINES += XLSX_NO_LIB

HEADERS += $$PWD/xlsxdocpropscore_p.h \
    $$PWD/xlsxdocpropsapp_p.h \
    $$PWD/xlsxrelationships_p.h \
    $$PWD/xlsxutility_p.h \
    $$PWD/xlsxsharedstrings_p.h \
    $$PWD/xlsxcontenttypes_p.h \
    $$PWD/xlsxtheme_p.h \
    $$PWD/xlsxformat.h \
    $$PWD/xlsxworkbook.h \
    $$PWD/xlsxstyles_p.h \
    $$PWD/xlsxabstractsheet.h \
    $$PWD/xlsxabstractsheet_p.h \
    $$PWD/xlsxworksheet.h \
    $$PWD/xlsxworksheet_p.h \
    $$PWD/xlsxchartsheet.h \
    $$PWD/xlsxchartsheet_p.h \
    $$PWD/xlsxzipwriter_p.h \
    $$PWD/xlsxworkbook_p.h \
    $$PWD/xlsxformat_p.h \
    $$PWD/xlsxglobal.h \
    $$PWD/xlsxdrawing_p.h \
    $$PWD/xlsxzipreader_p.h \
    $$PWD/xlsxdocument.h \
    $$PWD/xlsxdocument_p.h \
    $$PWD/xlsxcell.h \
    $$PWD/xlsxcell_p.h \
    $$PWD/xlsxdatavalidation.h \
    $$PWD/xlsxdatavalidation_p.h \
    $$PWD/xlsxcellreference.h \
    $$PWD/xlsxcellrange.h \
    $$PWD/xlsxrichstring_p.h \
    $$PWD/xlsxrichstring.h \
    $$PWD/xlsxconditionalformatting.h \
    $$PWD/xlsxconditionalformatting_p.h \
    $$PWD/xlsxcolor_p.h \
    $$PWD/xlsxnumformatparser_p.h \
    $$PWD/xlsxdrawinganchor_p.h \
    $$PWD/xlsxmediafile_p.h \
    $$PWD/xlsxabstractooxmlfile.h \
    $$PWD/xlsxabstractooxmlfile_p.h \
    $$PWD/xlsxchart.h \
    $$PWD/xlsxchart_p.h \
    $$PWD/xlsxsimpleooxmlfile_p.h \
    $$PWD/xlsxcellformula.h \
    $$PWD/xlsxcellformula_p.h

SOURCES += $$PWD/xlsxdocpropscore.cpp \
    $$PWD/xlsxdocpropsapp.cpp \
    $$PWD/xlsxrelationships.cpp \
    $$PWD/xlsxutility.cpp \
    $$PWD/xlsxsharedstrings.cpp \
    $$PWD/xlsxcontenttypes.cpp \
    $$PWD/xlsxtheme.cpp \
    $$PWD/xlsxformat.cpp \
    $$PWD/xlsxstyles.cpp \
    $$PWD/xlsxworkbook.cpp \
    $$PWD/xlsxabstractsheet.cpp \
    $$PWD/xlsxworksheet.cpp \
    $$PWD/xlsxchartsheet.cpp \
    $$PWD/xlsxzipwriter.cpp \
    $$PWD/xlsxdrawing.cpp \
    $$PWD/xlsxzipreader.cpp \
    $$PWD/xlsxdocument.cpp \
    $$PWD/xlsxcell.cpp \
    $$PWD/xlsxdatavalidation.cpp \
    $$PWD/xlsxcellreference.cpp \
    $$PWD/xlsxcellrange.cpp \
    $$PWD/xlsxrichstring.cpp \
    $$PWD/xlsxconditionalformatting.cpp \
    $$PWD/xlsxcolor.cpp \
    $$PWD/xlsxnumformatparser.cpp \
    $$PWD/xlsxdrawinganchor.cpp \
    $$PWD/xlsxmediafile.cpp \
    $$PWD/xlsxabstractooxmlfile.cpp \
    $$PWD/xlsxchart.cpp \
    $$PWD/xlsxsimpleooxmlfile.cpp \
    $$PWD/xlsxcellformula.cpp

