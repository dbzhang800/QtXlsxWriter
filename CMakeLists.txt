cmake_minimum_required(VERSION 3.2)

project(QtXlsxWriter)
add_definitions(-DQT_BUILD_XLSX_LIB)
set(BUILD_SHARED_LIBS TRUE)
set(CMAKE_AUTOMOC ON)

if (NOT CMAKE_BUILD_TYPE)
  set(CMAKE_BUILD_TYPE "Release")
endif()

file(
    GLOB
    QtXlsxWriter_SOURCE_FILES 
    ${CMAKE_CURRENT_SOURCE_DIR}/src/xlsx/*.cpp
    ${CMAKE_CURRENT_BINARY_DIR}/QtXlsxWriterTest_automoc.cpp
)

find_package(Qt5Core CONFIG REQUIRED)
find_package(Qt5Gui CONFIG REQUIRED)
find_package(Qt5Test CONFIG REQUIRED)

# hack to make private part of QtGui available
set(Qt5Gui_PRIVATE_INCLUDE_DIRS )
set(Qt5Core_PRIVATE_INCLUDE_DIRS )
foreach(_qt_gui_include_dir ${Qt5Gui_INCLUDE_DIRS})
  if(${_qt_gui_include_dir} MATCHES "include/QtGui")
	  list(APPEND Qt5Gui_PRIVATE_INCLUDE_DIRS "${_qt_gui_include_dir}/${Qt5Gui_VERSION}")
	  list(APPEND Qt5Gui_PRIVATE_INCLUDE_DIRS "${_qt_gui_include_dir}/${Qt5Gui_VERSION}/QtGui")
  endif()
  if(${_qt_gui_include_dir} MATCHES "include/QtCore")
	  list(APPEND Qt5Gui_PRIVATE_INCLUDE_DIRS "${_qt_gui_include_dir}/${Qt5Core_VERSION}")
	  list(APPEND Qt5Gui_PRIVATE_INCLUDE_DIRS "${_qt_gui_include_dir}/${Qt5Core_VERSION}/QtCore") 
  fi()
endforeach()

include_directories(
	${CMAKE_CURRENT_SOURCE_DIR}/src/xlsx/
	${Qt5Core_INCLUDE_DIRS} 
	${Qt5Gui_INCLUDE_DIRS}
	${Qt5Gui_PRIVATE_INCLUDE_DIRS} )

message("-- Using Qt private include in: ${Qt5Gui_PRIVATE_INCLUDE_DIRS}")

add_library(QtXlsxWriter SHARED "${QtXlsxWriter_SOURCE_FILES}")

# automatically add C++11 support with GCC
if(NOT MSVC)
    target_compile_features(QtXlsxWriter PRIVATE cxx_range_for)
endif()

set_target_properties(QtXlsxWriter PROPERTIES DEBUG_POSTFIX "d")
target_link_libraries(QtXlsxWriter Qt5::Gui Qt5::Core)

if(BUILD_TESTING)
  add_subdirectory(tests)
endif()

if(BUILD_EXAMPLES)
    add_custom_command(TARGET QtXlsxWriter POST_BUILD
        COMMAND ${CMAKE_COMMAND}
        -E copy_directory $<CONFIGURATION> ${CMAKE_CURRENT_BINARY_DIR}/examples/xlsx/$<CONFIGURATION>)
	add_subdirectory(examples/xlsx)
endif()

##
#
# QtxlsxwriterVersion.cmake creation
#
##
set(QtXlsxWriter_CONFIG_PATH ${CMAKE_INSTALL_PREFIX})
configure_file(QtXlsxWriterConfig.cmake.in QtXlsxWriterConfig.cmake @ONLY)

#####
#
# Installation configuration
#
#####
INSTALL(TARGETS QtXlsxWriter
  RUNTIME DESTINATION bin
  LIBRARY DESTINATION lib
  ARCHIVE DESTINATION lib
)

INSTALL(FILES  ${CMAKE_CURRENT_BINARY_DIR}/QtXlsxWriterConfig.cmake DESTINATION cmake)

SET(INCLUDE_FILES 
  src/xlsx/xlsxabstractooxmlfile.h
  src/xlsx/xlsxabstractsheet.h
  src/xlsx/xlsxcell.h
  src/xlsx/xlsxcellformula.h
  src/xlsx/xlsxcellrange.h
  src/xlsx/xlsxcellreference.h
  src/xlsx/xlsxchart.h
  src/xlsx/xlsxchartsheet.h
  src/xlsx/xlsxconditionalformatting.h
  src/xlsx/xlsxdatavalidation.h
  src/xlsx/xlsxdocument.h
  src/xlsx/xlsxformat.h
  src/xlsx/xlsxglobal.h
  src/xlsx/xlsxrichstring.h
  src/xlsx/xlsxworkbook.h
  src/xlsx/xlsxworksheet.h
)
INSTALL(FILES ${INCLUDE_FILES} DESTINATION include)

