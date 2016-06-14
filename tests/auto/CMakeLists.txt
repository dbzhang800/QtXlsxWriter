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

# Create a fake private/ include folder in
# the compilation tree
set( 
	QtXlsxWriter_PRIVATE_SOURCE_FILES 
    	${CMAKE_CURRENT_BINARY_DIR}/include)

if(DEBUG)
	message("Private include directory : ${QtXlsxWriter_PRIVATE_SOURCE_FILES}")
endif()

file(MAKE_DIRECTORY "${QtXlsxWriter_PRIVATE_SOURCE_FILES}/private")

# and copy all private includes in it
file(GLOB private_includes ${CMAKE_CURRENT_SOURCE_DIR}/../../src/xlsx/*_p.h)
foreach(private_include ${private_includes})
	if(DEBUG)
		message("Copying private include : ${private_include}") 
	endif()
	file(COPY ${private_include} DESTINATION ${QtXlsxWriter_PRIVATE_SOURCE_FILES}/private)
endforeach()


add_definitions(-DXLSX_TESTS)
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
	add_executable(${test}_testDriver ${TEST_SRC})
  	target_link_libraries(${test}_testDriver QtXlsxWriter)
	target_link_libraries(${test}_testDriver ${Qt5Core_LIBRARIES})
	target_link_libraries(${test}_testDriver ${Qt5Gui_LIBRARIES})
	target_link_libraries(${test}_testDriver ${Qt5Test_LIBRARIES})
	add_test(NAME ${test} COMMAND ${test}_testDriver -C $<CONFIGURATION>)
endforeach(test)
