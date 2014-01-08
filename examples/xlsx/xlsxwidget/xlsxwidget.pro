TARGET = xlsxwidget

#include(../../../src/xlsx/qtxlsx.pri)
QT+= xlsx xlsx-private widgets

SOURCES += main.cpp \
    xlsxsheetmodel.cpp

HEADERS += \
    xlsxsheetmodel.h \
    xlsxsheetmodel_p.h
