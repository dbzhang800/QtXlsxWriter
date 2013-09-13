QT       += testlib xlsx xlsx-private
CONFIG += testcase
DEFINES += XLSX_TEST

TARGET = tst_stylestest
CONFIG   += console
CONFIG   -= app_bundle

TEMPLATE = app


SOURCES += tst_stylestest.cpp
DEFINES += SRCDIR=\\\"$$PWD/\\\"
