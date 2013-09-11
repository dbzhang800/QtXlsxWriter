QT       += testlib xlsx xlsx-private
CONFIG += testcase
DEFINES += XLSX_TEST

TARGET = tst_propscoretest
CONFIG   += console
CONFIG   -= app_bundle

TEMPLATE = app


SOURCES += tst_propscoretest.cpp
DEFINES += SRCDIR=\\\"$$PWD/\\\"
