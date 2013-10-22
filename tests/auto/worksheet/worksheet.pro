#-------------------------------------------------
#
# Project created by QtCreator 2013-09-06T10:52:56
#
#-------------------------------------------------

QT       += testlib xlsx xlsx-private
CONFIG += testcase
DEFINES += XLSX_TEST

TARGET = tst_worksheet
CONFIG   += console
CONFIG   -= app_bundle

TEMPLATE = app

SOURCES += tst_worksheet.cpp
DEFINES += SRCDIR=\\\"$$PWD/\\\"
