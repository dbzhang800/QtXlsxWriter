TARGET = mergecells

#include(../../../src/xlsx/qtxlsx.pri)
QT += xlsx

TARGET = numberformat
CONFIG   += console
CONFIG   -= app_bundle

TEMPLATE = app

SOURCES += main.cpp
