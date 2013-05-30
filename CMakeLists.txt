cmake_minimum_required ( VERSION 2.6.4 )
set( LOCAL_PROJECT_NAME test )
project ( Superbuild_${LOCAL_PROJECT_NAME} )
# The version number.
set ( DMouse_VERSION_MAJOR 1 )
set ( DMouse_VERSION_MINOR 4 )

include( "${CMAKE_CURRENT_SOURCE_DIR}/Superbuild.cmake" )

set(EXTERNAL_PROJECT_BUILD_TYPE "Release" CACHE STRING "Default build type for support libraries")
option(USE_SYSTEM_OpenCV "Build using an externally defined version of OpenCV" OFF)


# build a CPack driven installer package
include ( InstallRequiredSystemLibraries )
# set ( CPACK_RESOURCE_FILE_LICENSE
#      "${CMAKE_CURRENT_SOURCE_DIR}/License.txt" ) ## TODO: include License.txt
set ( CPACK_PACKAGE_VERSION_MAJOR "${DMouse_VERSION_MAJOR}" )
set ( CPACK_PACKAGE_VERSION_MINOR "${DMouse_VERSION_MINOR}" )
include ( CPack )
