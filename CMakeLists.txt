cmake_minimum_required(VERSION 3.20)
project(uphide LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_BUILD_TYPE Release)

file(GLOB SRC "src/*.cpp")
list(APPEND SRC "src/uphide.rc")
add_executable(uphide ${SRC})
target_link_libraries(uphide ole32 oleaut32 uuid)
if(MSVC)
	set_target_properties(uphide PROPERTIES LINK_FLAGS "/SUBSYSTEM:CONSOLE /MANIFEST:NO" VS_CONFIGURATION_TYPE Application)
endif()
