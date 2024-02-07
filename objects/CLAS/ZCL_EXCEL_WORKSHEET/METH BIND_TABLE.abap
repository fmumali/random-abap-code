  METHOD bind_table.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmöcker,      (wi p)              2012-12-01
*              - ...
*          aligning code
*          message made to support multilinguality
*--------------------------------------------------------------------*
* issue #237   - Check if overlapping areas exist
*              - Alessandro Iannacci                        2012-12-01
* changes:     - Added raise if overlaps are detected
*--------------------------------------------------------------------*

    CONSTANTS:
      lc_top_left_column TYPE zexcel_cell_column_alpha VALUE 'A',
      lc_top_left_row    TYPE zexcel_cell_row VALUE 1,
      lc_no_currency     TYPE waers_curc VALUE IS INITIAL.

    DATA:
      lv_row_int              TYPE zexcel_cell_row,
      lv_first_row            TYPE zexcel_cell_row,
      lv_last_row             TYPE zexcel_cell_row,
      lv_column_int           TYPE zexcel_cell_column,
      lv_column_alpha         TYPE zexcel_cell_column_alpha,
      lt_field_catalog        TYPE zexcel_t_fieldcatalog,
      lv_id                   TYPE i,
      lv_formula              TYPE string,
      ls_settings             TYPE zexcel_s_table_settings,
      lo_table                TYPE REF TO zcl_excel_table,
      lv_value_lowercase      TYPE string,
      lv_syindex              TYPE c LENGTH 3,
      lo_iterator             TYPE REF TO zcl_excel_collection_iterator,
      lo_style_cond           TYPE REF TO zcl_excel_style_cond,
      lo_curtable             TYPE REF TO zcl_excel_table,
      lt_other_table_settings TYPE ty_table_settings.
    DATA: ls_column_formula TYPE mty_s_column_formula,
          lv_mincol         TYPE i.

    FIELD-SYMBOLS:
      <ls_field_catalog>        TYPE zexcel_s_fieldcatalog,
      <ls_field_catalog_custom> TYPE zexcel_s_fieldcatalog,
      <fs_table_line>           TYPE any,
      <fs_fldval>               TYPE any,
      <fs_fldval_currency>      TYPE waers.

    ls_settings = is_table_settings.

    IF ls_settings-top_left_column IS INITIAL.
      ls_settings-top_left_column = lc_top_left_column.
    ENDIF.

    IF ls_settings-table_style IS INITIAL.
      ls_settings-table_style = zcl_excel_table=>builtinstyle_medium2.
    ENDIF.

    IF ls_settings-top_left_row IS INITIAL.
      ls_settings-top_left_row = lc_top_left_row.
    ENDIF.

    IF it_field_catalog IS NOT SUPPLIED.
      lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = ip_table
                                                             ip_conv_exit_length = ip_conv_exit_length ).
    ELSE.
      lt_field_catalog = it_field_catalog.
    ENDIF.

    SORT lt_field_catalog BY position.

    calculate_table_bottom_right(
      EXPORTING
        ip_table         = ip_table
        it_field_catalog = lt_field_catalog
      CHANGING
        cs_settings      = ls_settings ).

* Check if overlapping areas exist

    lo_iterator = me->tables->get_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_curtable ?= lo_iterator->get_next( ).
      APPEND lo_curtable->settings TO lt_other_table_settings.
    ENDWHILE.

    check_table_overlapping(
        is_table_settings       = ls_settings
        it_other_table_settings = lt_other_table_settings ).

* Start filling the table

    CREATE OBJECT lo_table.
    lo_table->settings = ls_settings.
    lo_table->set_data( ir_data = ip_table ).
    lv_id = me->excel->get_next_table_id( ).
    lo_table->set_id( iv_id = lv_id ).

    me->tables->add( lo_table ).

    lv_column_int = zcl_excel_common=>convert_column2int( ls_settings-top_left_column ).
    lv_row_int = ls_settings-top_left_row.

    lt_field_catalog = normalize_column_heading_texts(
          iv_default_descr = iv_default_descr
          it_field_catalog = lt_field_catalog ).

* It is better to loop column by column (only visible column)
    LOOP AT lt_field_catalog ASSIGNING <ls_field_catalog> WHERE dynpfld EQ abap_true.

      lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column_int ).

      " First of all write column header
      IF <ls_field_catalog>-style_header IS NOT INITIAL.
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = <ls_field_catalog>-column_name
                      ip_style  = <ls_field_catalog>-style_header ).
      ELSE.
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = <ls_field_catalog>-column_name ).
      ENDIF.

      me->set_table_reference( ip_column    = lv_column_int
                               ip_row       = lv_row_int
                               ir_table     = lo_table
                               ip_fieldname = <ls_field_catalog>-fieldname
                               ip_header    = abap_true ).

      IF <ls_field_catalog>-column_formula IS NOT INITIAL.
        ls_column_formula-id                     = lines( column_formulas ) + 1.
        ls_column_formula-column                 = lv_column_int.
        ls_column_formula-formula                = <ls_field_catalog>-column_formula.
        ls_column_formula-table_top_left_row     = lo_table->settings-top_left_row.
        ls_column_formula-table_bottom_right_row = lo_table->settings-bottom_right_row.
        ls_column_formula-table_left_column_int  = lv_mincol.
        ls_column_formula-table_right_column_int = zcl_excel_common=>convert_column2int( lo_table->settings-bottom_right_column ).
        INSERT ls_column_formula INTO TABLE column_formulas.
      ENDIF.

      ADD 1 TO lv_row_int.
      LOOP AT ip_table ASSIGNING <fs_table_line>.

        ASSIGN COMPONENT <ls_field_catalog>-fieldname OF STRUCTURE <fs_table_line> TO <fs_fldval>.

        " issue #290 Add formula support in table
        IF <ls_field_catalog>-formula EQ abap_true.
          IF <ls_field_catalog>-style IS NOT INITIAL.
            IF <ls_field_catalog>-abap_type IS NOT INITIAL.
              me->set_cell( ip_column   = lv_column_alpha
                          ip_row      = lv_row_int
                          ip_formula  = <fs_fldval>
                          ip_abap_type = <ls_field_catalog>-abap_type
                          ip_style    = <ls_field_catalog>-style ).
            ELSE.
              me->set_cell( ip_column   = lv_column_alpha
                            ip_row      = lv_row_int
                            ip_formula  = <fs_fldval>
                            ip_style    = <ls_field_catalog>-style ).
            ENDIF.
          ELSEIF <ls_field_catalog>-abap_type IS NOT INITIAL.
            me->set_cell( ip_column   = lv_column_alpha
                          ip_row      = lv_row_int
                          ip_formula  = <fs_fldval>
                          ip_abap_type = <ls_field_catalog>-abap_type ).
          ELSE.
            me->set_cell( ip_column   = lv_column_alpha
                          ip_row      = lv_row_int
                          ip_formula  = <fs_fldval> ).
          ENDIF.
        ELSEIF <ls_field_catalog>-column_formula IS NOT INITIAL.
          " Column formulas
          IF <ls_field_catalog>-style IS NOT INITIAL.
            IF <ls_field_catalog>-abap_type IS NOT INITIAL.
              me->set_cell( ip_column            = lv_column_alpha
                            ip_row               = lv_row_int
                            ip_column_formula_id = ls_column_formula-id
                            ip_abap_type         = <ls_field_catalog>-abap_type
                            ip_style             = <ls_field_catalog>-style ).
            ELSE.
              me->set_cell( ip_column            = lv_column_alpha
                            ip_row               = lv_row_int
                            ip_column_formula_id = ls_column_formula-id
                            ip_style             = <ls_field_catalog>-style ).
            ENDIF.
          ELSEIF <ls_field_catalog>-abap_type IS NOT INITIAL.
            me->set_cell( ip_column             = lv_column_alpha
                          ip_row                = lv_row_int
                          ip_column_formula_id  = ls_column_formula-id
                          ip_abap_type          = <ls_field_catalog>-abap_type ).
          ELSE.
            me->set_cell( ip_column            = lv_column_alpha
                          ip_row               = lv_row_int
                          ip_column_formula_id = ls_column_formula-id ).
          ENDIF.
        ELSE.
          IF <ls_field_catalog>-currency_column IS INITIAL OR ip_conv_curr_amt_ext = abap_false.
            ASSIGN lc_no_currency TO <fs_fldval_currency>.
          ELSE.
            ASSIGN COMPONENT <ls_field_catalog>-currency_column OF STRUCTURE <fs_table_line> TO <fs_fldval_currency>.
          ENDIF.

          IF <ls_field_catalog>-style IS NOT INITIAL.
            IF <ls_field_catalog>-abap_type IS NOT INITIAL.
              me->set_cell( ip_column           = lv_column_alpha
                            ip_row              = lv_row_int
                            ip_value            = <fs_fldval>
                            ip_abap_type        = <ls_field_catalog>-abap_type
                            ip_currency         = <fs_fldval_currency>
                            ip_style            = <ls_field_catalog>-style
                            ip_conv_exit_length = ip_conv_exit_length ).
            ELSE.
              me->set_cell( ip_column = lv_column_alpha
                            ip_row    = lv_row_int
                            ip_value  = <fs_fldval>
                            ip_currency = <fs_fldval_currency>
                            ip_style  = <ls_field_catalog>-style
                            ip_conv_exit_length = ip_conv_exit_length ).
            ENDIF.
          ELSE.
            IF <ls_field_catalog>-abap_type IS NOT INITIAL.
              me->set_cell( ip_column = lv_column_alpha
                          ip_row    = lv_row_int
                          ip_abap_type = <ls_field_catalog>-abap_type
                          ip_currency  = <fs_fldval_currency>
                          ip_value  = <fs_fldval>
                          ip_conv_exit_length = ip_conv_exit_length ).
            ELSE.
              me->set_cell( ip_column = lv_column_alpha
                            ip_row    = lv_row_int
                            ip_currency = <fs_fldval_currency>
                            ip_value  = <fs_fldval>
                            ip_conv_exit_length = ip_conv_exit_length ).
            ENDIF.
          ENDIF.
        ENDIF.
        ADD 1 TO lv_row_int.

      ENDLOOP.
      IF sy-subrc <> 0 AND iv_no_line_if_empty = abap_false. "create empty row if table has no data
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = space ).
        ADD 1 TO lv_row_int.
      ENDIF.

*--------------------------------------------------------------------*
      " totals
*--------------------------------------------------------------------*
      IF <ls_field_catalog>-totals_function IS NOT INITIAL.
        lv_formula = lo_table->get_totals_formula( ip_column = <ls_field_catalog>-column_name ip_function = <ls_field_catalog>-totals_function ).
        IF <ls_field_catalog>-style_total IS NOT INITIAL.
          me->set_cell( ip_column   = lv_column_alpha
                        ip_row      = lv_row_int
                        ip_formula  = lv_formula
                        ip_style    = <ls_field_catalog>-style_total ).
        ELSE.
          me->set_cell( ip_column   = lv_column_alpha
                        ip_row      = lv_row_int
                        ip_formula  = lv_formula ).
        ENDIF.
      ENDIF.

      lv_row_int = ls_settings-top_left_row.
      ADD 1 TO lv_column_int.

*--------------------------------------------------------------------*
      " conditional formatting
*--------------------------------------------------------------------*
      IF <ls_field_catalog>-style_cond IS NOT INITIAL.
        lv_first_row    = ls_settings-top_left_row + 1. " +1 to exclude header
        lv_last_row     = ls_settings-bottom_right_row.
        lo_style_cond = me->get_style_cond( <ls_field_catalog>-style_cond ).
        lo_style_cond->set_range( ip_start_column  = lv_column_alpha
                                  ip_start_row     = lv_first_row
                                  ip_stop_column   = lv_column_alpha
                                  ip_stop_row      = lv_last_row ).
      ENDIF.

    ENDLOOP.

*--------------------------------------------------------------------*
    " Set field catalog
*--------------------------------------------------------------------*
    lo_table->fieldcat = lt_field_catalog[].

    es_table_settings = ls_settings.
    es_table_settings-bottom_right_column = lv_column_alpha.
    " >> Issue #291
    IF ip_table IS INITIAL.
      es_table_settings-bottom_right_row    = ls_settings-top_left_row + 2.           "Last rows
    ELSE.
      es_table_settings-bottom_right_row    = ls_settings-bottom_right_row + 1. "Last rows
    ENDIF.
    " << Issue #291

  ENDMETHOD.                    "BIND_TABLE