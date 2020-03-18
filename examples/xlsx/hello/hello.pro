TARGET = hello

#include(../../../src/xlsx/qtxlsx.pri)
QT+=xlsx

CONFIG   += console
CONFIG   -= app_bundle

SOURCES += main.cpp

# install
target.path = $$[QT_INSTALL_EXAMPLES]/xlsx/hello
INSTALLS += target
