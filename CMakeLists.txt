cmake_minimum_required (VERSION 2.8)
project (libbhohelper)

set(CMAKE_CONFIGURATION_TYPES Debug Release )
set(CMAKE_CONFIGURATION_TYPES "${CMAKE_CONFIGURATION_TYPES}" CACHE STRING "Reset the configurations to those we need" FORCE)

#Variables describing project layout
set(${PROJECT_NAME}_INC_DIR "${${PROJECT_NAME}_SOURCE_DIR}/libbhohelper"  CACHE PATH "Path to ${PROJECT_NAME} include directory")
set(${PROJECT_NAME}_SRC_DIR "${${PROJECT_NAME}_SOURCE_DIR}/libbhohelper")

include_directories("${${PROJECT_NAME}_INC_DIR}")

set(${PROJECT_NAME}_HEADERS
  "${${PROJECT_NAME}_INC_DIR}/libbhohelper.h"
  "${${PROJECT_NAME}_INC_DIR}/stdafx.h"
  )

set(${PROJECT_NAME}_SOURCES
  "${${PROJECT_NAME}_SRC_DIR}/AtlFnInfos.cpp"
  "${${PROJECT_NAME}_SRC_DIR}/IDispatchHelper.cpp"
  "${${PROJECT_NAME}_SRC_DIR}/ComLibrary.cpp"
  "${${PROJECT_NAME}_SRC_DIR}/stdafx.cpp"
  #  "${${PROJECT_NAME}_SRC_DIR}/uri.cpp"
  )
################# MAIN TARGET SETUP #########################

add_library(${PROJECT_NAME} STATIC ${${PROJECT_NAME}_HEADERS} ${${PROJECT_NAME}_SOURCES})

# set UNICODE
set_target_properties(${PROJECT_NAME} PROPERTIES COMPILE_DEFINITIONS _UNICODE,UNICODE)

# _WIN64
if (${CMAKE_CL_64})
  add_definitions(-DWIN64 -D_WIN64)
endif (${CMAKE_CL_64})

# exclude default libraries
set_target_properties(${PROJECT_NAME} PROPERTIES STATIC_LIBRARY_FLAGS /NODEFAULTLIB)

set(${PROJECT_NAME}_LIB_DIR "${${PROJECT_NAME}_BINARY_DIR}"  CACHE PATH "Path to ${PROJECT_NAME} lib directory")


message(STATUS "Configuring ${PROJECT_NAME}")
