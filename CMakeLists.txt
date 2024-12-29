cmake_minimum_required(VERSION 3.5)

project(Qt_Cmake_MySQL LANGUAGES CXX)

set(CMAKE_INCLUDE_CURRENT_DIR ON)

set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# QtCreator supports the following variables for Android, which are identical to qmake Android variables.
# Check http://doc.qt.io/qt-5/deployment-android.html for more information.
# They need to be set before the find_package(Qt5 ...) call.

#if(ANDROID)
#    set(ANDROID_PACKAGE_SOURCE_DIR "${CMAKE_CURRENT_SOURCE_DIR}/android")
#    if (ANDROID_ABI STREQUAL "armeabi-v7a")
#        set(ANDROID_EXTRA_LIBS
#            ${CMAKE_CURRENT_SOURCE_DIR}/path/to/libcrypto.so
#            ${CMAKE_CURRENT_SOURCE_DIR}/path/to/libssl.so)
#    endif()
#endif()

find_package(Qt5 COMPONENTS Widgets REQUIRED)

add_subdirectory(QXlsx)  #添加

if(ANDROID)
  add_library(Qt_Cmake_MySQL SHARED
    main.cpp
    mainwindow.cpp
    mainwindow.h
    mainwindow.ui
  )
else()
  add_executable(Qt_Cmake_MySQL
    main.cpp
    mainwindow.cpp
    mainwindow.h
    mainwindow.ui
  )
endif()


find_package(Qt5 COMPONENTS Widgets Sql AxContainer REQUIRED)  # 添加 Qt 模块
target_link_libraries(Qt_Cmake_MySQL PRIVATE Qt5::Sql)
# mysql版本  Ver 8.0.39 for Win64 on x86_64 (MySQL Community Server - GPL)
# Workbench版本 Version 8.0. 38 build 4270059 CE (64 bits) Community

target_link_libraries(Qt_Cmake_MySQL PRIVATE Qt5::Widgets QXlsx::QXlsx Qt5::AxContainer)
