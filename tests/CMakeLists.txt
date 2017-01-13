#
# create a test-enabled library before creating tests
# themselves (export more functions than regular DLL)
#

# Create a fake private/ include folder 
# in the compilation tree to allow the
# #include "private/..." instruction to
# work
set( 
	QtXlsxWriter_PRIVATE_SOURCE_FILES 
    	${CMAKE_CURRENT_BINARY_DIR}/include)

if(DEBUG)
	message("Private include directory : ${QtXlsxWriter_PRIVATE_SOURCE_FILES}")
endif()
file(MAKE_DIRECTORY "${QtXlsxWriter_PRIVATE_SOURCE_FILES}/private")

# Copy all private includes in the created folder
file(GLOB private_includes ${CMAKE_CURRENT_SOURCE_DIR}/../src/xlsx/*_p.h)
foreach(private_include ${private_includes})
	if(DEBUG)
		message("Copying private include : ${private_include}") 
	endif()
	file(COPY ${private_include} DESTINATION ${QtXlsxWriter_PRIVATE_SOURCE_FILES}/private)
endforeach()

set(CMAKE_AUTOMOC ON)

file(
  GLOB
  QtXlsxWriterTest_SOURCE_FILES
  ${CMAKE_CURRENT_SOURCE_DIR}/../src/xlsx/*.cpp
  ${CMAKE_CURRENT_BINARY_DIR}/QtXlsxWriter_automoc.cpp
)

include_directories(
  ${CMAKE_CURRENT_BINARY_DIR}
  ${Qt5Core_INCLUDE_DIRS} 
  ${Qt5Gui_INCLUDE_DIRS}
  ${Qt5Gui_PRIVATE_INCLUDE_DIRS}
)
  
add_definitions(-DQT_BUILD_XLSX_LIB)
add_definitions(-DXLSX_TEST)
add_library(QtXlsxWriterTest SHARED "${QtXlsxWriterTest_SOURCE_FILES}")

# automatically add C++11 support with GCC
if(NOT MSVC)
    target_compile_features(QtXlsxWriterTest PRIVATE cxx_range_for)
endif()

 
set_target_properties(QtXlsxWriterTest PROPERTIES DEBUG_POSTFIX "d")
  
target_link_libraries(QtXlsxWriterTest ${Qt5Core_LIBRARIES})
target_link_libraries(QtXlsxWriterTest ${Qt5Gui_LIBRARIES})

add_custom_command(TARGET QtXlsxWriterTest POST_BUILD
                     COMMAND ${CMAKE_COMMAND}
                     -E copy_directory $<CONFIGURATION> auto/$<CONFIGURATION>)

add_subdirectory(auto)
