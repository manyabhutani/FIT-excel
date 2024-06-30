cmake_minimum_required(VERSION 3.25)
project(excel)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall -pedantic -g -fsanitize=address")

add_executable(excel test.cpp)
target_link_directories(excel PUBLIC ./arm64-darwin23-clang)
target_link_libraries(excel PUBLIC expression_parser)

