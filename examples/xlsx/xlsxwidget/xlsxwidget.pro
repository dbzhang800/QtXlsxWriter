TARGET = xlsxwidget
QT += widgets

#include(../../../src/xlsx/qtxlsx.pri)
QT+= xlsx

SOURCES += main.cpp \
    xlsxsheetmodel.cpp

HEADERS += \
    xlsxsheetmodel.h \
    xlsxsheetmodel_p.h
