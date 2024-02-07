  METHOD validate_range.

    DATA: lv_range_name       TYPE string,
          lv_range_address    TYPE string,
          lv_range_start      TYPE string,
          lv_range_stop       TYPE string,
          lv_range_sheet      TYPE string,
          lv_tmp_value        TYPE string,
          lt_cell_coord_parts TYPE TABLE OF string,
          lv_cell_coord_start TYPE string,
          lv_cell_coord_stop  TYPE string,
          lv_column_start     TYPE zexcel_cell_column_alpha,
          lv_column_end       TYPE zexcel_cell_column_alpha,
          lv_row_start        TYPE zexcel_cell_row,
          lv_row_end          TYPE zexcel_cell_row.

    FIELD-SYMBOLS: <ls_range> TYPE ts_range.


    lv_range_name = io_range->name.
    TRANSLATE lv_range_name TO UPPER CASE.
    lv_range_address = io_range->get_value( ).

    SPLIT lv_range_address AT '!' INTO lv_range_sheet lv_tmp_value.

    SPLIT lv_tmp_value AT ':' INTO lv_range_start lv_range_stop.

    SPLIT lv_range_start AT '$' INTO TABLE lt_cell_coord_parts.

    IF lines( lt_cell_coord_parts ) > 2.
      TRY.
          zcl_excel_common=>convert_range2column_a_row(
            EXPORTING
              i_range        = lv_range_address
            IMPORTING
              e_column_start = lv_column_start
              e_column_end   = lv_column_end
              e_row_start    = lv_row_start
              e_row_end      = lv_row_end
          ).
        CATCH zcx_excel.    "
          RETURN.
      ENDTRY.
      IF lv_column_start = 'A' AND lv_column_end = 'XFD'.
        lv_cell_coord_start = |{ lv_row_start }|.
        lv_cell_coord_stop = |{ lv_row_end }|.
        CLEAR lt_cell_coord_parts.
        CLEAR lv_range_stop.
      ELSE.
        RETURN.
      ENDIF.
    ENDIF.

    IF lines( lt_cell_coord_parts ) >= 2.
      READ TABLE lt_cell_coord_parts INTO lv_cell_coord_start INDEX 2.
    ENDIF.

    IF lv_cell_coord_start CO '0123456789'.
      APPEND INITIAL LINE TO mt_range ASSIGNING <ls_range>.
      <ls_range>-sheet = lv_range_sheet.
      <ls_range>-name = lv_range_name.
      <ls_range>-start = lv_cell_coord_start.

      SPLIT lv_range_stop AT '$' INTO TABLE lt_cell_coord_parts.
      READ TABLE lt_cell_coord_parts INTO lv_cell_coord_stop INDEX 2.
      <ls_range>-stop = lv_cell_coord_stop.

      <ls_range>-length = <ls_range>-stop - <ls_range>-start + 1.

    ENDIF.

  ENDMETHOD.