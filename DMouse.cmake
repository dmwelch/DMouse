include(${CMAKE_CURRENT_LIST_DIR}/Common.cmake)

find_package( OpenCV REQUIRED )
if( OpenCV_FOUND )
  include( ${OpenCV_USE_FILE})
endif()

add_executable( DMouse dmouse.cxx )

# configure a header file to pass some of the CMake settings
# to the source code
configure_file (
  "${PROJECT_SOURCE_DIR}/DMouseConfig.h.in"
  "${PROJECT_BINARY_DIR}/DMouseConfig.h"
 )

# add the binary tree to the search path for include files
# so that we will find DMouseConfig.h
include_directories( "${PROJECT_BINARY_DIR}" )

### HACK: libxl
include_directories( "${CMAKE_CURRENT_SOURCE_DIR}/../libxl-3.4.1.4/include_cpp/libxl.h" )
### END HACK