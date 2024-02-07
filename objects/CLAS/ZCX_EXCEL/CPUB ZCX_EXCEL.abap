CLASS zcx_excel DEFINITION
  PUBLIC
  INHERITING FROM cx_static_check
  CREATE PUBLIC .

*"* public components of class ZCX_EXCEL
*"* do not include other source files here!!!
*"* protected components of class ZCX_EXCEL
*"* do not include other source files here!!!
*"* protected components of class ZCX_EXCEL
*"* do not include other source files here!!!
  PUBLIC SECTION.

    CONSTANTS zcx_excel TYPE sotr_conc VALUE '028C0ED2B5601ED78EB6F3368B1E4F9B' ##NO_TEXT.
    DATA error TYPE string .
    DATA syst_at_raise TYPE syst .

    METHODS constructor
      IMPORTING
        !textid        LIKE textid OPTIONAL
        !previous      LIKE previous OPTIONAL
        !error         TYPE string OPTIONAL
        !syst_at_raise TYPE syst OPTIONAL .
    CLASS-METHODS raise_text
      IMPORTING
        !iv_text TYPE clike
      RAISING
        zcx_excel .
    CLASS-METHODS raise_symsg
      RAISING
        zcx_excel .

    METHODS if_message~get_longtext
        REDEFINITION .
    METHODS if_message~get_text
        REDEFINITION .