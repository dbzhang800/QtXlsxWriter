QT       += testlib xlsx # xlsx-private
CONFIG += testcase
DEFINES += XLSX_TEST

TARGET = tst_format
CONFIG   += console
CONFIG   -= app_bundle

TEMPLATE = app

SOURCES += tst_formattest.cpp
DEFINES += SRCDIR=\\\"$$PWD/\\\"
