class-pool .
*"* class pool for class ZCL_EXCEL_COMMON

*"* local type definitions
include ZCL_EXCEL_COMMON==============ccdef.

*"* class ZCL_EXCEL_COMMON definition
*"* public declarations
  include ZCL_EXCEL_COMMON==============cu.
*"* protected declarations
  include ZCL_EXCEL_COMMON==============co.
*"* private declarations
  include ZCL_EXCEL_COMMON==============ci.
endclass. "ZCL_EXCEL_COMMON definition

*"* macro definitions
include ZCL_EXCEL_COMMON==============ccmac.
*"* local class implementation
include ZCL_EXCEL_COMMON==============ccimp.

*"* test class
include ZCL_EXCEL_COMMON==============ccau.

class ZCL_EXCEL_COMMON implementation.
*"* method's implementations
  include methods.
endclass. "ZCL_EXCEL_COMMON implementation
