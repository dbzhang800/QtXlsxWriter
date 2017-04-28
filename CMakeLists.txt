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

find_package(Qt5 5.5 REQUIRED Core Gui Test)
include_directories(
	${CMAKE_CURRENT_SOURCE_DIR}/src/xlsx/
	${Qt5Core_INCLUDE_DIRS} 
	${Qt5Gui_INCLUDE_DIRS}
	${Qt5Gui_PRIVATE_INCLUDE_DIRS} )

add_library(QtXlsxWriter SHARED "${QtXlsxWriter_SOURCE_FILES}")

# automatically add C++11 support with GCC
if(NOT MSVC)
    target_compile_features(QtXlsxWriter PRIVATE cxx_range_for)
endif()

set_target_properties(QtXlsxWriter PROPERTIES DEBUG_POSTFIX "d")
target_link_libraries(QtXlsxWriter ${Qt5Core_LIBRARIES})
target_link_libraries(QtXlsxWriter ${Qt5Gui_LIBRARIES})

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

INSTALL(FILES  ${CMAKE_CURRENT_BINARY_DIR}/QtXlsxWriterConfig.cmake DESTINATION lib/cmake/${PROJECT_NAME})

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

