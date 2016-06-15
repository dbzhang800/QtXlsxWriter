cmake_minimum_required(VERSION 3.2)

project(QtXlsxWriter)
add_definitions(-DQT_BUILD_XLSX_LIB)
SET (BUILD_SHARED_LIBS TRUE)

set(CMAKE_AUTOMOC ON)

file(
	GLOB
	QtXlsxWriter_SOURCE_FILES 
	${CMAKE_CURRENT_SOURCE_DIR}/src/xlsx/*.cpp
)

find_package(Qt5 5.5 REQUIRED Core Gui Test)
include_directories(
	${QtXlsxWriter_SOURCE_FILES} 
	${Qt5Core_INCLUDE_DIRS} 
	${Qt5Gui_INCLUDE_DIRS}
	${Qt5Gui_PRIVATE_INCLUDE_DIRS} )

add_library(QtXlsxWriter SHARED "${QtXlsxWriter_SOURCE_FILES}")

set_target_properties(QtXlsxWriter PROPERTIES DEBUG_POSTFIX "d")
target_link_libraries(QtXlsxWriter ${Qt5Core_LIBRARIES})
target_link_libraries(QtXlsxWriter ${Qt5Gui_LIBRARIES})

if(BUILD_TESTS)

add_custom_command(TARGET QtXlsxWriter POST_BUILD
  COMMAND ${CMAKE_COMMAND} 
  -E copy_directory $<CONFIGURATION>/ ${CMAKE_CURRENT_BINARY_DIR}/tests/auto/$<CONFIGURATION>/)
	add_subdirectory(tests)
endif()

if(BUILD_EXAMPLES)
	add_subdirectory(examples)
endif()
