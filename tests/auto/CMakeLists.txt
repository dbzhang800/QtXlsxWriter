set(TESTS 
	cellreference  
	document 
	format 
	propsapp 
	propscore 
	relationships 
	richstring 
	sharedstrings 
	styles 
	utility 
	worksheet 
	xlsxconditionalformatting 
	zipreader)
  
enable_testing()

# Compile each defined test and
# declare it for ctest tool
foreach(test IN ITEMS ${TESTS})
	if(DEBUG)
		message("configure test '${test}'")
	endif()
	file( GLOB TEST_SRC ${test}/*.cpp)
	include_directories(
		${CMAKE_CURRENT_SOURCE_DIR}/../../src/xlsx/
		${Qt5Core_INCLUDE_DIRS} 
		${Qt5Gui_INCLUDE_DIRS}
    ${QtXlsxWriter_PRIVATE_SOURCE_FILES}
		${CMAKE_CURRENT_BINARY_DIR} # .moc files
	)
  
  add_definitions(-DXLSX_TEST)
  remove_definitions(-DQT_BUILD_XLSX_LIB)

	add_executable(${test}_testDriver ${TEST_SRC} )

# automatically add C++11 support with GCC
if(NOT MSVC)
    target_compile_features(${test}_testDriver PRIVATE cxx_range_for)
endif()


  target_link_libraries(${test}_testDriver ${Qt5Core_LIBRARIES})
	target_link_libraries(${test}_testDriver ${Qt5Gui_LIBRARIES})
	target_link_libraries(${test}_testDriver ${Qt5Test_LIBRARIES})
  target_link_libraries(${test}_testDriver QtXlsxWriterTest)

  add_test(NAME ${test} COMMAND ${test}_testDriver)

endforeach(test)
