cmake_minimum_required(VERSION 3.2)

project(QtXlsxWriter)
add_definitions(-DXLSX_NO_LIB)
SET (BUILD_SHARED_LIBS  TRUE) 

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

target_link_libraries(QtXlsxWriter ${Qt5Core_LIBRARIES})
target_link_libraries(QtXlsxWriter ${Qt5Gui_LIBRARIES})

if(BUILD_TESTS)
	add_subdirectory(tests)
endif()

if(BUILD_EXAMPLES)
	add_subdirectory(examples)
endif()
