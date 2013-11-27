QT       += testlib xlsx xlsx-private
CONFIG += testcase
DEFINES += XLSX_TEST

TARGET = tst_conditionalformattingtest
CONFIG   += console
CONFIG   -= app_bundle

TEMPLATE = app

SOURCES += tst_conditionalformattingtest.cpp
DEFINES += SRCDIR=\\\"$$PWD/\\\"
