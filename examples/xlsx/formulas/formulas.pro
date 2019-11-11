TARGET = hello

#include(../../../src/xlsx/qtxlsx.pri)
QT+=xlsx

CONFIG   += console
CONFIG   += install_ok
CONFIG   -= app_bundle

SOURCES += main.cpp
