project(QtXlsx)

include_directories(${Qt5Gui_PRIVATE_INCLUDE_DIRS})
set(CMAKE_AUTOMOC ON)

FILE(GLOB PROJECT_SRCS src/xlsx/*.cpp src/xlsx/*.h)

add_library(${PROJECT_NAME}_lib STATIC ${PROJECT_SRCS})

target_link_libraries(${PROJECT_NAME}_lib
                      Qt5::Core Qt5::Widgets)
