  METHOD load_worksheet.

    super->load_worksheet( EXPORTING ip_path      = ip_path
                                     io_worksheet = io_worksheet ).

    DATA: node      TYPE REF TO if_ixml_element,
          worksheet TYPE REF TO if_ixml_document,
          sheetpr   TYPE t_sheetpr.

    " Read Workbook codeName
    worksheet = me->get_ixml_from_zip_archive( ip_path ).
    node ?=  worksheet->find_from_name_ns( name = 'sheetPr' uri = namespace-main ).
    IF node IS BOUND.

      fill_struct_from_attributes( EXPORTING ip_element   = node
                                   CHANGING  cp_structure = sheetpr  ).
      IF sheetpr-codename IS NOT INITIAL.
        io_worksheet->zif_excel_sheet_vba_project~set_codename_pr( sheetpr-codename ).
      ENDIF.
    ENDIF.
  ENDMETHOD.