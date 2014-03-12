QT       += testlib #xlsx # xlsx-private
CONFIG += testcase
DEFINES += XLSX_TEST

TARGET = tst_xmlspacetest
CONFIG   += console
CONFIG   -= app_bundle

TEMPLATE = app


SOURCES += tst_xmlspacetest.cpp
DEFINES += SRCDIR=\\\"$$PWD/\\\"
