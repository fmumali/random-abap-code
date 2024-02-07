class-pool .
*"* class pool for class ZCL_EXCEL_WORKSHEET

*"* local type definitions
include ZCL_EXCEL_WORKSHEET===========ccdef.

*"* class ZCL_EXCEL_WORKSHEET definition
*"* public declarations
  include ZCL_EXCEL_WORKSHEET===========cu.
*"* protected declarations
  include ZCL_EXCEL_WORKSHEET===========co.
*"* private declarations
  include ZCL_EXCEL_WORKSHEET===========ci.
endclass. "ZCL_EXCEL_WORKSHEET definition

*"* macro definitions
include ZCL_EXCEL_WORKSHEET===========ccmac.
*"* local class implementation
include ZCL_EXCEL_WORKSHEET===========ccimp.

*"* test class
include ZCL_EXCEL_WORKSHEET===========ccau.

class ZCL_EXCEL_WORKSHEET implementation.
*"* method's implementations
  include methods.
endclass. "ZCL_EXCEL_WORKSHEET implementation
