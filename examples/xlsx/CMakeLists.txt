find_package(Qt5 5.5 REQUIRED Core Gui Test Widgets)

set(EXAMPLES
    calendar
    chart
    chartsheet
    conditionalformatting
    datavalidation
    definename
    demo
    documentproperty
    extractdata
    formulas
    hello
    hyperlinks
    image
    mergecells
    numberformat
    richtext
    rowcolumn
    style
    worksheetoperations
    xlsxwidget)

# Compile each defined example 
foreach(example IN ITEMS ${EXAMPLES})
	if(DEBUG)
		message("Create example '${example}'")
	endif()
	file( GLOB TEST_SRC ${example}/*.cpp)
	include_directories(
		${CMAKE_CURRENT_SOURCE_DIR}/../../src/xlsx/
		${Qt5Core_INCLUDE_DIRS} 
		${Qt5Gui_INCLUDE_DIRS}
		${Qt5Widgets_INCLUDE_DIRS}
        ${QtXlsxWriter_PRIVATE_SOURCE_FILES}
		${CMAKE_CURRENT_BINARY_DIR} # .moc files
	)
  
  remove_definitions(-DQT_BUILD_XLSX_LIB)

  add_executable(${example} ${TEST_SRC} )

  target_link_libraries(${example} ${Qt5Widgets_LIBRARIES})
  target_link_libraries(${example} ${Qt5Core_LIBRARIES})
  target_link_libraries(${example} ${Qt5Gui_LIBRARIES})
  target_link_libraries(${example} ${Qt5Test_LIBRARIES})
  target_link_libraries(${example} QtXlsxWriter)

endforeach(example)
