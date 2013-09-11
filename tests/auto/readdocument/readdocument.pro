QT       += testlib xlsx xlsx-private
CONFIG += testcase
DEFINES += XLSX_TEST

TARGET = tst_readdocumenttest
CONFIG   += console
CONFIG   -= app_bundle

TEMPLATE = app


SOURCES += tst_readdocumenttest.cpp
DEFINES += SRCDIR=\\\"$$PWD/\\\"
