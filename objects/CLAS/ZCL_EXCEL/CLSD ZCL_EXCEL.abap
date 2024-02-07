class-pool .
*"* class pool for class ZCL_EXCEL

*"* local type definitions
include ZCL_EXCEL=====================ccdef.

*"* class ZCL_EXCEL definition
*"* public declarations
  include ZCL_EXCEL=====================cu.
*"* protected declarations
  include ZCL_EXCEL=====================co.
*"* private declarations
  include ZCL_EXCEL=====================ci.
endclass. "ZCL_EXCEL definition

*"* macro definitions
include ZCL_EXCEL=====================ccmac.
*"* local class implementation
include ZCL_EXCEL=====================ccimp.

*"* test class
include ZCL_EXCEL=====================ccau.

class ZCL_EXCEL implementation.
*"* method's implementations
  include methods.
endclass. "ZCL_EXCEL implementation
