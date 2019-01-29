TARGET = numberformat

#include(../../../src/xlsx/qtxlsx.pri)
QT += xlsx

CONFIG   += console
CONFIG   += install_ok
CONFIG   -= app_bundle

TEMPLATE = app

SOURCES += main.cpp
