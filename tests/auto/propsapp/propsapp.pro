QT       += testlib xlsx xlsx-private
CONFIG += testcase
DEFINES += XLSX_TEST

TARGET = tst_docpropsapptest
CONFIG   += console
CONFIG   -= app_bundle

TEMPLATE = app


SOURCES += tst_docpropsapptest.cpp
DEFINES += SRCDIR=\\\"$$PWD/\\\"
