QT       += testlib xlsx xlsx-private
CONFIG += testcase
DEFINES += XLSX_TEST

TARGET = tst_richstringtest
CONFIG   += console
CONFIG   -= app_bundle

TEMPLATE = app

SOURCES += tst_richstringtest.cpp
DEFINES += SRCDIR=\\\"$$PWD/\\\"
