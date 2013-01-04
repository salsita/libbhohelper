cmake_minimum_required (VERSION 2.8)
project (libbhohelper)

set(${PROJECT_NAME}_INC_DIRS "${${PROJECT_NAME}_SOURCE_DIR}/libbhohelper")
set(${PROJECT_NAME}_SRC_DIR "${${PROJECT_NAME}_SOURCE_DIR}/libbhohelper")

include_directories("${${PROJECT_NAME}_INC_DIRS}")

set(${PROJECT_NAME}_HEADERS
	"${${PROJECT_NAME}_INC_DIRS}/libbhohelper.h"
	"${${PROJECT_NAME}_INC_DIRS}/stdafx.h"
	)

set(${PROJECT_NAME}_SOURCES 
	"${${PROJECT_NAME}_SRC_DIR}/AtlFnInfos.cpp"
	"${${PROJECT_NAME}_SRC_DIR}/IDispatchHelper.cpp"
	"${${PROJECT_NAME}_SRC_DIR}/stdafx.cpp"
	#	"${${PROJECT_NAME}_SRC_DIR}/uri.cpp"
	)

add_library(${PROJECT_NAME} STATIC ${${PROJECT_NAME}_HEADERS} ${${PROJECT_NAME}_SOURCES})

# set UNICODE
set_target_properties(${PROJECT_NAME} PROPERTIES COMPILE_DEFINITIONS _UNICODE,UNICODE)

# exclude default libraries
set_target_properties(${PROJECT_NAME} PROPERTIES STATIC_LIBRARY_FLAGS /NODEFAULTLIB)

set(${PROJECT_NAME}_INC_DIRS ${${PROJECT_NAME}_INC_DIRS} PARENT_SCOPE)

#set(${PROJECT_NAME}_ARCHIVE_DIRS "${${PROJECT_NAME}_BINARY_DIR}/Debug" PARENT_SCOPE)
#get_cmake_property(_variableNames VARIABLES)
#foreach (_variableName ${_variableNames})
#    message(STATUS "${_variableName}=${${_variableName}}")
#endforeach()

message(STATUS "Generated ${PROJECT_NAME}")
