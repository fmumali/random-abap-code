  METHOD change_cell_style.

    DATA: changer TYPE REF TO zif_excel_style_changer,
          column  TYPE zexcel_cell_column,
          row     TYPE zexcel_cell_row.

    normalize_columnrow_parameter( EXPORTING ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   IMPORTING ep_column    = column
                                             ep_row       = row ).

    changer = zcl_excel_style_changer=>create( excel = excel ).


    IF ip_complete IS SUPPLIED.
      IF ip_xcomplete IS NOT SUPPLIED.
        zcx_excel=>raise_text( 'Complete styleinfo has to be supplied with corresponding X-field' ).
      ENDIF.
      changer->set_complete( ip_complete = ip_complete ip_xcomplete = ip_xcomplete ).
    ENDIF.



    IF ip_font IS SUPPLIED.
      IF ip_xfont IS SUPPLIED.
        changer->set_complete_font( ip_font = ip_font ip_xfont = ip_xfont ).
      ELSE.
        changer->set_complete_font( ip_font = ip_font ).
      ENDIF.
    ENDIF.

    IF ip_fill IS SUPPLIED.
      IF ip_xfill IS SUPPLIED.
        changer->set_complete_fill( ip_fill = ip_fill ip_xfill = ip_xfill ).
      ELSE.
        changer->set_complete_fill( ip_fill = ip_fill ).
      ENDIF.
    ENDIF.


    IF ip_borders IS SUPPLIED.
      IF ip_xborders IS SUPPLIED.
        changer->set_complete_borders( ip_borders = ip_borders ip_xborders = ip_xborders ).
      ELSE.
        changer->set_complete_borders( ip_borders = ip_borders ).
      ENDIF.
    ENDIF.

    IF ip_alignment IS SUPPLIED.
      IF ip_xalignment IS SUPPLIED.
        changer->set_complete_alignment( ip_alignment = ip_alignment ip_xalignment = ip_xalignment ).
      ELSE.
        changer->set_complete_alignment( ip_alignment = ip_alignment ).
      ENDIF.
    ENDIF.

    IF ip_protection IS SUPPLIED.
      IF ip_xprotection IS SUPPLIED.
        changer->set_complete_protection( ip_protection = ip_protection ip_xprotection = ip_xprotection ).
      ELSE.
        changer->set_complete_protection( ip_protection = ip_protection ).
      ENDIF.
    ENDIF.


    IF ip_borders_allborders IS SUPPLIED.
      IF ip_xborders_allborders IS SUPPLIED.
        changer->set_complete_borders_all( ip_borders_allborders = ip_borders_allborders ip_xborders_allborders = ip_xborders_allborders ).
      ELSE.
        changer->set_complete_borders_all( ip_borders_allborders = ip_borders_allborders ).
      ENDIF.
    ENDIF.

    IF ip_borders_diagonal IS SUPPLIED.
      IF ip_xborders_diagonal IS SUPPLIED.
        changer->set_complete_borders_diagonal( ip_borders_diagonal = ip_borders_diagonal ip_xborders_diagonal = ip_xborders_diagonal ).
      ELSE.
        changer->set_complete_borders_diagonal( ip_borders_diagonal = ip_borders_diagonal ).
      ENDIF.
    ENDIF.

    IF ip_borders_down IS SUPPLIED.
      IF ip_xborders_down IS SUPPLIED.
        changer->set_complete_borders_down( ip_borders_down = ip_borders_down ip_xborders_down = ip_xborders_down ).
      ELSE.
        changer->set_complete_borders_down( ip_borders_down = ip_borders_down ).
      ENDIF.
    ENDIF.

    IF ip_borders_left IS SUPPLIED.
      IF ip_xborders_left IS SUPPLIED.
        changer->set_complete_borders_left( ip_borders_left = ip_borders_left ip_xborders_left = ip_xborders_left ).
      ELSE.
        changer->set_complete_borders_left( ip_borders_left = ip_borders_left ).
      ENDIF.
    ENDIF.

    IF ip_borders_right IS SUPPLIED.
      IF ip_xborders_right IS SUPPLIED.
        changer->set_complete_borders_right( ip_borders_right = ip_borders_right ip_xborders_right = ip_xborders_right ).
      ELSE.
        changer->set_complete_borders_right( ip_borders_right = ip_borders_right ).
      ENDIF.
    ENDIF.

    IF ip_borders_top IS SUPPLIED.
      IF ip_xborders_top IS SUPPLIED.
        changer->set_complete_borders_top( ip_borders_top = ip_borders_top ip_xborders_top = ip_xborders_top ).
      ELSE.
        changer->set_complete_borders_top( ip_borders_top = ip_borders_top ).
      ENDIF.
    ENDIF.

    IF ip_number_format_format_code IS SUPPLIED.
      changer->set_number_format( ip_number_format_format_code ).
    ENDIF.
    IF ip_font_bold IS SUPPLIED.
      changer->set_font_bold( ip_font_bold ).
    ENDIF.
    IF ip_font_color IS SUPPLIED.
      changer->set_font_color( ip_font_color ).
    ENDIF.
    IF ip_font_color_rgb IS SUPPLIED.
      changer->set_font_color_rgb( ip_font_color_rgb ).
    ENDIF.
    IF ip_font_color_indexed IS SUPPLIED.
      changer->set_font_color_indexed( ip_font_color_indexed ).
    ENDIF.
    IF ip_font_color_theme IS SUPPLIED.
      changer->set_font_color_theme( ip_font_color_theme ).
    ENDIF.
    IF ip_font_color_tint IS SUPPLIED.
      changer->set_font_color_tint( ip_font_color_tint ).
    ENDIF.

    IF ip_font_family IS SUPPLIED.
      changer->set_font_family( ip_font_family ).
    ENDIF.
    IF ip_font_italic IS SUPPLIED.
      changer->set_font_italic( ip_font_italic ).
    ENDIF.
    IF ip_font_name IS SUPPLIED.
      changer->set_font_name( ip_font_name ).
    ENDIF.
    IF ip_font_scheme IS SUPPLIED.
      changer->set_font_scheme( ip_font_scheme ).
    ENDIF.
    IF ip_font_size IS SUPPLIED.
      changer->set_font_size( ip_font_size ).
    ENDIF.
    IF ip_font_strikethrough IS SUPPLIED.
      changer->set_font_strikethrough( ip_font_strikethrough ).
    ENDIF.
    IF ip_font_underline IS SUPPLIED.
      changer->set_font_underline( ip_font_underline ).
    ENDIF.
    IF ip_font_underline_mode IS SUPPLIED.
      changer->set_font_underline_mode( ip_font_underline_mode ).
    ENDIF.

    IF ip_fill_filltype IS SUPPLIED.
      changer->set_fill_filltype( ip_fill_filltype ).
    ENDIF.
    IF ip_fill_rotation IS SUPPLIED.
      changer->set_fill_rotation( ip_fill_rotation ).
    ENDIF.
    IF ip_fill_fgcolor IS SUPPLIED.
      changer->set_fill_fgcolor( ip_fill_fgcolor ).
    ENDIF.
    IF ip_fill_fgcolor_rgb IS SUPPLIED.
      changer->set_fill_fgcolor_rgb( ip_fill_fgcolor_rgb ).
    ENDIF.
    IF ip_fill_fgcolor_indexed IS SUPPLIED.
      changer->set_fill_fgcolor_indexed( ip_fill_fgcolor_indexed ).
    ENDIF.
    IF ip_fill_fgcolor_theme IS SUPPLIED.
      changer->set_fill_fgcolor_theme( ip_fill_fgcolor_theme ).
    ENDIF.
    IF ip_fill_fgcolor_tint IS SUPPLIED.
      changer->set_fill_fgcolor_tint( ip_fill_fgcolor_tint ).
    ENDIF.

    IF ip_fill_bgcolor IS SUPPLIED.
      changer->set_fill_bgcolor( ip_fill_bgcolor ).
    ENDIF.
    IF ip_fill_bgcolor_rgb IS SUPPLIED.
      changer->set_fill_bgcolor_rgb( ip_fill_bgcolor_rgb ).
    ENDIF.
    IF ip_fill_bgcolor_indexed IS SUPPLIED.
      changer->set_fill_bgcolor_indexed( ip_fill_bgcolor_indexed ).
    ENDIF.
    IF ip_fill_bgcolor_theme IS SUPPLIED.
      changer->set_fill_bgcolor_theme( ip_fill_bgcolor_theme ).
    ENDIF.
    IF ip_fill_bgcolor_tint IS SUPPLIED.
      changer->set_fill_bgcolor_tint( ip_fill_bgcolor_tint ).
    ENDIF.

    IF ip_fill_gradtype_type IS SUPPLIED.
      changer->set_fill_gradtype_type( ip_fill_gradtype_type ).
    ENDIF.
    IF ip_fill_gradtype_degree IS SUPPLIED.
      changer->set_fill_gradtype_degree( ip_fill_gradtype_degree ).
    ENDIF.
    IF ip_fill_gradtype_bottom IS SUPPLIED.
      changer->set_fill_gradtype_bottom( ip_fill_gradtype_bottom ).
    ENDIF.
    IF ip_fill_gradtype_left IS SUPPLIED.
      changer->set_fill_gradtype_left( ip_fill_gradtype_left ).
    ENDIF.
    IF ip_fill_gradtype_top IS SUPPLIED.
      changer->set_fill_gradtype_top( ip_fill_gradtype_top ).
    ENDIF.
    IF ip_fill_gradtype_right IS SUPPLIED.
      changer->set_fill_gradtype_right( ip_fill_gradtype_right ).
    ENDIF.
    IF ip_fill_gradtype_position1 IS SUPPLIED.
      changer->set_fill_gradtype_position1( ip_fill_gradtype_position1 ).
    ENDIF.
    IF ip_fill_gradtype_position2 IS SUPPLIED.
      changer->set_fill_gradtype_position2( ip_fill_gradtype_position2 ).
    ENDIF.
    IF ip_fill_gradtype_position3 IS SUPPLIED.
      changer->set_fill_gradtype_position3( ip_fill_gradtype_position3 ).
    ENDIF.



    IF ip_borders_diagonal_mode IS SUPPLIED.
      changer->set_borders_diagonal_mode( ip_borders_diagonal_mode ).
    ENDIF.
    IF ip_alignment_horizontal IS SUPPLIED.
      changer->set_alignment_horizontal( ip_alignment_horizontal ).
    ENDIF.
    IF ip_alignment_vertical IS SUPPLIED.
      changer->set_alignment_vertical( ip_alignment_vertical ).
    ENDIF.
    IF ip_alignment_textrotation IS SUPPLIED.
      changer->set_alignment_textrotation( ip_alignment_textrotation ).
    ENDIF.
    IF ip_alignment_wraptext IS SUPPLIED.
      changer->set_alignment_wraptext( ip_alignment_wraptext ).
    ENDIF.
    IF ip_alignment_shrinktofit IS SUPPLIED.
      changer->set_alignment_shrinktofit( ip_alignment_shrinktofit ).
    ENDIF.
    IF ip_alignment_indent IS SUPPLIED.
      changer->set_alignment_indent( ip_alignment_indent ).
    ENDIF.
    IF ip_protection_hidden IS SUPPLIED.
      changer->set_protection_hidden( ip_protection_hidden ).
    ENDIF.
    IF ip_protection_locked IS SUPPLIED.
      changer->set_protection_locked( ip_protection_locked ).
    ENDIF.

    IF ip_borders_allborders_style IS SUPPLIED.
      changer->set_borders_allborders_style( ip_borders_allborders_style ).
    ENDIF.
    IF ip_borders_allborders_color IS SUPPLIED.
      changer->set_borders_allborders_color( ip_borders_allborders_color ).
    ENDIF.
    IF ip_borders_allbo_color_rgb IS SUPPLIED.
      changer->set_borders_allbo_color_rgb( ip_borders_allbo_color_rgb ).
    ENDIF.
    IF ip_borders_allbo_color_indexed IS SUPPLIED.
      changer->set_borders_allbo_color_indexe( ip_borders_allbo_color_indexed ).
    ENDIF.
    IF ip_borders_allbo_color_theme IS SUPPLIED.
      changer->set_borders_allbo_color_theme( ip_borders_allbo_color_theme ).
    ENDIF.
    IF ip_borders_allbo_color_tint IS SUPPLIED.
      changer->set_borders_allbo_color_tint( ip_borders_allbo_color_tint ).
    ENDIF.

    IF ip_borders_diagonal_style IS SUPPLIED.
      changer->set_borders_diagonal_style( ip_borders_diagonal_style ).
    ENDIF.
    IF ip_borders_diagonal_color IS SUPPLIED.
      changer->set_borders_diagonal_color( ip_borders_diagonal_color ).
    ENDIF.
    IF ip_borders_diagonal_color_rgb IS SUPPLIED.
      changer->set_borders_diagonal_color_rgb( ip_borders_diagonal_color_rgb ).
    ENDIF.
    IF ip_borders_diagonal_color_inde IS SUPPLIED.
      changer->set_borders_diagonal_color_ind( ip_borders_diagonal_color_inde ).
    ENDIF.
    IF ip_borders_diagonal_color_them IS SUPPLIED.
      changer->set_borders_diagonal_color_the( ip_borders_diagonal_color_them ).
    ENDIF.
    IF ip_borders_diagonal_color_tint IS SUPPLIED.
      changer->set_borders_diagonal_color_tin( ip_borders_diagonal_color_tint ).
    ENDIF.

    IF ip_borders_down_style IS SUPPLIED.
      changer->set_borders_down_style( ip_borders_down_style ).
    ENDIF.
    IF ip_borders_down_color IS SUPPLIED.
      changer->set_borders_down_color( ip_borders_down_color ).
    ENDIF.
    IF ip_borders_down_color_rgb IS SUPPLIED.
      changer->set_borders_down_color_rgb( ip_borders_down_color_rgb ).
    ENDIF.
    IF ip_borders_down_color_indexed IS SUPPLIED.
      changer->set_borders_down_color_indexed( ip_borders_down_color_indexed ).
    ENDIF.
    IF ip_borders_down_color_theme IS SUPPLIED.
      changer->set_borders_down_color_theme( ip_borders_down_color_theme ).
    ENDIF.
    IF ip_borders_down_color_tint IS SUPPLIED.
      changer->set_borders_down_color_tint( ip_borders_down_color_tint ).
    ENDIF.

    IF ip_borders_left_style IS SUPPLIED.
      changer->set_borders_left_style( ip_borders_left_style ).
    ENDIF.
    IF ip_borders_left_color IS SUPPLIED.
      changer->set_borders_left_color( ip_borders_left_color ).
    ENDIF.
    IF ip_borders_left_color_rgb IS SUPPLIED.
      changer->set_borders_left_color_rgb( ip_borders_left_color_rgb ).
    ENDIF.
    IF ip_borders_left_color_indexed IS SUPPLIED.
      changer->set_borders_left_color_indexed( ip_borders_left_color_indexed ).
    ENDIF.
    IF ip_borders_left_color_theme IS SUPPLIED.
      changer->set_borders_left_color_theme( ip_borders_left_color_theme ).
    ENDIF.
    IF ip_borders_left_color_tint IS SUPPLIED.
      changer->set_borders_left_color_tint( ip_borders_left_color_tint ).
    ENDIF.

    IF ip_borders_right_style IS SUPPLIED.
      changer->set_borders_right_style( ip_borders_right_style ).
    ENDIF.
    IF ip_borders_right_color IS SUPPLIED.
      changer->set_borders_right_color( ip_borders_right_color ).
    ENDIF.
    IF ip_borders_right_color_rgb IS SUPPLIED.
      changer->set_borders_right_color_rgb( ip_borders_right_color_rgb ).
    ENDIF.
    IF ip_borders_right_color_indexed IS SUPPLIED.
      changer->set_borders_right_color_indexe( ip_borders_right_color_indexed ).
    ENDIF.
    IF ip_borders_right_color_theme IS SUPPLIED.
      changer->set_borders_right_color_theme( ip_borders_right_color_theme ).
    ENDIF.
    IF ip_borders_right_color_tint IS SUPPLIED.
      changer->set_borders_right_color_tint( ip_borders_right_color_tint ).
    ENDIF.

    IF ip_borders_top_style IS SUPPLIED.
      changer->set_borders_top_style( ip_borders_top_style ).
    ENDIF.
    IF ip_borders_top_color IS SUPPLIED.
      changer->set_borders_top_color( ip_borders_top_color ).
    ENDIF.
    IF ip_borders_top_color_rgb IS SUPPLIED.
      changer->set_borders_top_color_rgb( ip_borders_top_color_rgb ).
    ENDIF.
    IF ip_borders_top_color_indexed IS SUPPLIED.
      changer->set_borders_top_color_indexed( ip_borders_top_color_indexed ).
    ENDIF.
    IF ip_borders_top_color_theme IS SUPPLIED.
      changer->set_borders_top_color_theme( ip_borders_top_color_theme ).
    ENDIF.
    IF ip_borders_top_color_tint IS SUPPLIED.
      changer->set_borders_top_color_tint( ip_borders_top_color_tint ).
    ENDIF.


    ep_guid = changer->apply( ip_worksheet = me
                              ip_column    = column
                              ip_row       = row ).


  ENDMETHOD.                    "CHANGE_CELL_STYLE