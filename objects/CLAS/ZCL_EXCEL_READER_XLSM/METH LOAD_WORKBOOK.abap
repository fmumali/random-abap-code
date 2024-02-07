  METHOD load_workbook.
    super->load_workbook( EXPORTING iv_workbook_full_filename = iv_workbook_full_filename
                                    io_excel                  = io_excel ).

    CONSTANTS: lc_vba_project  TYPE string VALUE 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'.

    DATA: rels_workbook_path TYPE string,
          rels_workbook      TYPE REF TO if_ixml_document,
          path               TYPE string,
          node               TYPE REF TO if_ixml_element,
          workbook           TYPE REF TO if_ixml_document,
          stripped_name      TYPE chkfile,
          dirname            TYPE string,
          relationship       TYPE t_relationship,
          fileversion        TYPE t_fileversion,
          workbookpr         TYPE t_workbookpr.

    CALL FUNCTION 'TRINT_SPLIT_FILE_AND_PATH'
      EXPORTING
        full_name     = iv_workbook_full_filename
      IMPORTING
        stripped_name = stripped_name
        file_path     = dirname.

    " Read Workbook Relationships
    CONCATENATE dirname '_rels/' stripped_name '.rels'
      INTO rels_workbook_path.

    rels_workbook = me->get_ixml_from_zip_archive( rels_workbook_path ).

    node ?= rels_workbook->find_from_name_ns( name = 'Relationship' uri = namespace-relationships ).
    WHILE node IS BOUND.
      me->fill_struct_from_attributes( EXPORTING ip_element = node CHANGING cp_structure = relationship ).

      CASE relationship-type.
        WHEN lc_vba_project.
          " Read VBA  binary
          CONCATENATE dirname relationship-target INTO path.
          me->load_vbaproject( ip_path  = path
                               ip_excel = io_excel ).
        WHEN OTHERS.
      ENDCASE.

      node ?= node->get_next( ).
    ENDWHILE.

    " Read Workbook codeName
    workbook = me->get_ixml_from_zip_archive( iv_workbook_full_filename ).
    node ?=  workbook->find_from_name_ns( name = 'fileVersion' uri = namespace-main ).
    IF node IS BOUND.

      fill_struct_from_attributes( EXPORTING ip_element   = node
                                   CHANGING  cp_structure = fileversion  ).

      io_excel->zif_excel_book_vba_project~set_codename( fileversion-codename ).
    ENDIF.

    " Read Workbook codeName
    workbook = me->get_ixml_from_zip_archive( iv_workbook_full_filename ).
    node ?=  workbook->find_from_name_ns( name = 'workbookPr' uri = namespace-main ).
    IF node IS BOUND.

      fill_struct_from_attributes( EXPORTING ip_element   = node
                                   CHANGING  cp_structure = workbookpr  ).

      io_excel->zif_excel_book_vba_project~set_codename_pr( workbookpr-codename ).
    ENDIF.

  ENDMETHOD.