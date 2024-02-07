  METHOD zif_excel_reader~load.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Map Document Properties to ZCL_EXCEL
*--------------------------------------------------------------------*

    CONSTANTS: lcv_core_properties TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
               lcv_office_document TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'.

    DATA: lo_rels         TYPE REF TO if_ixml_document,
          lo_node         TYPE REF TO if_ixml_element,
          ls_relationship TYPE t_relationship.

*--------------------------------------------------------------------*
* §1  Create EXCEL-Object we want to return to caller

* §2  We need to read the the file "\\_rels\.rels" because it tells
*     us where in this folder structure the data for the workbook
*     is located in the xlsx zip-archive
*
*     The xlsx Zip-archive has generally the following folder structure:
*       <root> |
*              |-->  _rels
*              |-->  doc_Props
*              |-->  xl |
*                       |-->  _rels
*                       |-->  theme
*                       |-->  worksheets

* §3  Extracting from this the path&file where the workbook is located
*     Following is an example how this file could be set up
*        <?xml version="1.0" encoding="UTF-8" standalone="true"?>
*        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
*            <Relationship Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Id="rId3"/>
*            <Relationship Target="docProps/core.xml" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Id="rId2"/>
*            <Relationship Target="xl/workbook.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Id="rId1"/>
*        </Relationships>
*--------------------------------------------------------------------*

    CLEAR mt_dxf_styles.
    CLEAR mt_ref_formulae.
    CLEAR shared_strings.
    CLEAR styles.

*--------------------------------------------------------------------*
* §1  Create EXCEL-Object we want to return to caller
*--------------------------------------------------------------------*
    IF iv_zcl_excel_classname IS INITIAL.
      CREATE OBJECT r_excel.
    ELSE.
      CREATE OBJECT r_excel TYPE (iv_zcl_excel_classname).
    ENDIF.

    zip = create_zip_archive( i_xlsx_binary = i_excel2007
                              i_use_alternate_zip = i_use_alternate_zip ).

*--------------------------------------------------------------------*
* §2  Get file in folderstructure
*--------------------------------------------------------------------*
    lo_rels = get_ixml_from_zip_archive( '_rels/.rels' ).

*--------------------------------------------------------------------*
* §3  Cycle through the Relationship Tags and use the ones we need
*--------------------------------------------------------------------*
    lo_node ?= lo_rels->find_from_name_ns( name = 'Relationship' uri = namespace-relationships ). "#EC NOTEXT
    WHILE lo_node IS BOUND.

      fill_struct_from_attributes( EXPORTING
                                     ip_element   = lo_node
                                   CHANGING
                                     cp_structure = ls_relationship ).
      CASE ls_relationship-type.

        WHEN lcv_office_document.
*--------------------------------------------------------------------*
* Parse workbook - main part here
*--------------------------------------------------------------------*
          load_workbook( iv_workbook_full_filename  = ls_relationship-target
                         io_excel                   = r_excel ).

        WHEN lcv_core_properties.
          " 2do§1   Map Document Properties to ZCL_EXCEL

        WHEN OTHERS.

      ENDCASE.
      lo_node ?= lo_node->get_next( ).

    ENDWHILE.


  ENDMETHOD.