QT       += testlib xlsx # xlsx-private
CONFIG += testcase
DEFINES += XLSX_TEST

TARGET = tst_document
CONFIG   += console
CONFIG   -= app_bundle

TEMPLATE = app


SOURCES += tst_documenttest.cpp
DEFINES += SRCDIR=\\\"$$PWD/\\\"
