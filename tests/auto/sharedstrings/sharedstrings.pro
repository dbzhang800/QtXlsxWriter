QT       += testlib xlsx xlsx-private
CONFIG += testcase
DEFINES += XLSX_TEST

TARGET = tst_sharedstringstest
CONFIG   += console
CONFIG   -= app_bundle

TEMPLATE = app


SOURCES += tst_sharedstringstest.cpp
DEFINES += SRCDIR=\\\"$$PWD/\\\"
