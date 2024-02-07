  METHOD get_cell.

    DATA: lv_column        TYPE zexcel_cell_column,
          lv_row           TYPE zexcel_cell_row,
          ls_sheet_content TYPE zexcel_s_cell_data.

    normalize_columnrow_parameter( EXPORTING ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   IMPORTING ep_column    = lv_column
                                             ep_row       = lv_row ).

    READ TABLE sheet_content INTO ls_sheet_content WITH TABLE KEY cell_row     = lv_row
                                                                  cell_column  = lv_column.

    ep_rc       = sy-subrc.
    ep_value    = ls_sheet_content-cell_value.
    ep_guid     = ls_sheet_content-cell_style.       " issue 139 - added this to be used for columnwidth calculation
    ep_formula  = ls_sheet_content-cell_formula.
    IF et_rtf IS SUPPLIED AND ls_sheet_content-rtf_tab IS NOT INITIAL.
      et_rtf = ls_sheet_content-rtf_tab.
    ENDIF.

    " Addition to solve issue #120, contribution by Stefan Schmöcker
    DATA: style_iterator TYPE REF TO zcl_excel_collection_iterator,
          style          TYPE REF TO zcl_excel_style.
    IF ep_style IS SUPPLIED.
      CLEAR ep_style.
      style_iterator = me->excel->get_styles_iterator( ).
      WHILE style_iterator->has_next( ) = abap_true.
        style ?= style_iterator->get_next( ).
        IF style->get_guid( ) = ls_sheet_content-cell_style.
          ep_style = style.
          EXIT.
        ENDIF.
      ENDWHILE.
    ENDIF.
  ENDMETHOD.                    "GET_CELL