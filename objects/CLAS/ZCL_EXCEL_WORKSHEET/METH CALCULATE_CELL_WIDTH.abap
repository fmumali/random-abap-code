  METHOD calculate_cell_width.
*--------------------------------------------------------------------*
* issue #293   - Roberto Bianco
*              - Christian Assig                            2014-03-14
*
* changes: - Calculate widths using SAPscript font metrics
*            (transaction SE73)
*          - Calculate the width of dates
*          - Add additional width for auto filter buttons
*          - Add cell padding to simulate Excel behavior
*--------------------------------------------------------------------*

    DATA: ld_cell_value                TYPE zexcel_cell_value,
          ld_style_guid                TYPE zexcel_cell_style,
          ls_stylemapping              TYPE zexcel_s_stylemapping,
          lo_table_object              TYPE REF TO object,
          lo_table                     TYPE REF TO zcl_excel_table,
          ld_table_top_left_column     TYPE zexcel_cell_column,
          ld_table_bottom_right_column TYPE zexcel_cell_column,
          ld_flag_contains_auto_filter TYPE abap_bool VALUE abap_false,
          ld_flag_bold                 TYPE abap_bool VALUE abap_false,
          ld_flag_italic               TYPE abap_bool VALUE abap_false,
          ld_date                      TYPE d,
          ld_date_char                 TYPE c LENGTH 50,
          ld_font_height               TYPE tdfontsize VALUE zcl_excel_font=>lc_default_font_height,
          ld_font_name                 TYPE zexcel_style_font_name VALUE zcl_excel_font=>lc_default_font_name.

    " Determine cell content and cell style
    me->get_cell( EXPORTING ip_column = ip_column
                            ip_row    = ip_row
                  IMPORTING ep_value  = ld_cell_value
                            ep_guid   = ld_style_guid ).

    " ABAP2XLSX uses tables to define areas containing headers and
    " auto-filters. Find out if the current cell is in the header
    " of one of these tables.
    LOOP AT me->tables->collection INTO lo_table_object.
      " Downcast: OBJECT -> ZCL_EXCEL_TABLE
      lo_table ?= lo_table_object.

      " Convert column letters to corresponding integer values
      ld_table_top_left_column =
        zcl_excel_common=>convert_column2int(
          lo_table->settings-top_left_column ).

      ld_table_bottom_right_column =
        zcl_excel_common=>convert_column2int(
          lo_table->settings-bottom_right_column ).

      " Is the current cell part of the table header?
      IF ip_column BETWEEN ld_table_top_left_column AND
                           ld_table_bottom_right_column AND
         ip_row    EQ lo_table->settings-top_left_row.
        " Current cell is part of the table header
        " -> Assume that an auto filter is present and that the font is
        "    bold
        ld_flag_contains_auto_filter = abap_true.
        ld_flag_bold = abap_true.
      ENDIF.
    ENDLOOP.

    " If a style GUID is present, read style attributes
    IF ld_style_guid IS NOT INITIAL.
      TRY.
          " Read style attributes
          ls_stylemapping = me->excel->get_style_to_guid( ld_style_guid ).

          " If the current cell contains the default date format,
          " convert the cell value to a date and calculate its length
          IF ls_stylemapping-complete_style-number_format-format_code =
             zcl_excel_style_number_format=>c_format_date_std.

            " Convert excel date to ABAP date
            ld_date =
              zcl_excel_common=>excel_string_to_date( ld_cell_value ).

            " Format ABAP date using user's formatting settings
            WRITE ld_date TO ld_date_char.

            " Remember the formatted date to calculate the cell size
            ld_cell_value = ld_date_char.

          ENDIF.

          " Read the font size and convert it to the font height
          " used by SAPscript (multiplication by 10)
          IF ls_stylemapping-complete_stylex-font-size = abap_true.
            ld_font_height = ls_stylemapping-complete_style-font-size * 10.
          ENDIF.

          " If set, remember the font name
          IF ls_stylemapping-complete_stylex-font-name = abap_true.
            ld_font_name = ls_stylemapping-complete_style-font-name.
          ENDIF.

          " If set, remember whether font is bold and italic.
          IF ls_stylemapping-complete_stylex-font-bold = abap_true.
            ld_flag_bold = ls_stylemapping-complete_style-font-bold.
          ENDIF.

          IF ls_stylemapping-complete_stylex-font-italic = abap_true.
            ld_flag_italic = ls_stylemapping-complete_style-font-italic.
          ENDIF.

        CATCH zcx_excel.                                "#EC NO_HANDLER
          " Style GUID is present, but style was not found
          " Continue with default values

      ENDTRY.
    ENDIF.

    ep_width = zcl_excel_font=>calculate_text_width(
      iv_font_name   = ld_font_name
      iv_font_height = ld_font_height
      iv_flag_bold   = ld_flag_bold
      iv_flag_italic = ld_flag_italic
      iv_cell_value  = ld_cell_value ).

    " If the current cell contains an auto filter, make it a bit wider.
    " The size used by the auto filter button does not depend on the font
    " size.
    IF ld_flag_contains_auto_filter = abap_true.
      ADD 2 TO ep_width.
    ENDIF.

  ENDMETHOD.