class-pool .
*"* class pool for class ZCL_ITAB_TO_EXCEL

*"* local type definitions
include ZCL_ITAB_TO_EXCEL=============ccdef.

*"* class ZCL_ITAB_TO_EXCEL definition
*"* public declarations
  include ZCL_ITAB_TO_EXCEL=============cu.
*"* protected declarations
  include ZCL_ITAB_TO_EXCEL=============co.
*"* private declarations
  include ZCL_ITAB_TO_EXCEL=============ci.
endclass. "ZCL_ITAB_TO_EXCEL definition

*"* macro definitions
include ZCL_ITAB_TO_EXCEL=============ccmac.
*"* local class implementation
include ZCL_ITAB_TO_EXCEL=============ccimp.

*"* test class
include ZCL_ITAB_TO_EXCEL=============ccau.

class ZCL_ITAB_TO_EXCEL implementation.
*"* method's implementations
  include methods.
endclass. "ZCL_ITAB_TO_EXCEL implementation
