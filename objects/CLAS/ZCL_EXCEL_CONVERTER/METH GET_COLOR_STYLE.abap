  METHOD get_color_style.
    DATA: ls_colors       TYPE zexcel_s_converter_col,
          ls_color_styles TYPE ts_color_styles,
          lo_style        TYPE REF TO zcl_excel_style.

    r_style = i_style. " Default we change nothing

    IF wt_colors IS NOT INITIAL.
* Full line has color
      READ TABLE wt_colors INTO ls_colors WITH KEY rownumber   = i_row
                                                   columnname  = space.
      IF sy-subrc = 0.
        READ TABLE wt_color_styles INTO ls_color_styles WITH KEY guid_old  = i_style
                                                                 fontcolor = ls_colors-fontcolor
                                                                 fillcolor = ls_colors-fillcolor.
        IF sy-subrc = 0.
          r_style = ls_color_styles-style_new->get_guid( ).
        ELSE.
          lo_style = create_color_style( i_style          = i_style
                                         is_colors        = ls_colors ) .
          r_style = lo_style->get_guid( ) .
          ls_color_styles-guid_old  = i_style.
          ls_color_styles-fontcolor = ls_colors-fontcolor.
          ls_color_styles-fillcolor = ls_colors-fillcolor.
          ls_color_styles-style_new = lo_style.
          INSERT ls_color_styles INTO TABLE wt_color_styles.
        ENDIF.
      ELSE.
* Only field has color
        READ TABLE wt_colors INTO ls_colors WITH KEY rownumber   = i_row
                                                     columnname  = i_fieldname.
        IF sy-subrc = 0.
          READ TABLE wt_color_styles INTO ls_color_styles WITH KEY guid_old  = i_style
                                                                   fontcolor = ls_colors-fontcolor
                                                                   fillcolor = ls_colors-fillcolor.
          IF sy-subrc = 0.
            r_style = ls_color_styles-style_new->get_guid( ).
          ELSE.
            lo_style = create_color_style( i_style          = i_style
                                           is_colors        = ls_colors ) .
            ls_color_styles-guid_old  = i_style.
            ls_color_styles-fontcolor = ls_colors-fontcolor.
            ls_color_styles-fillcolor = ls_colors-fillcolor.
            ls_color_styles-style_new = lo_style.
            INSERT ls_color_styles INTO TABLE wt_color_styles.
            r_style = ls_color_styles-style_new->get_guid( ).
          ENDIF.
        ELSE.
          r_style = i_style.
        ENDIF.
      ENDIF.
    ELSE.
      r_style = i_style.
    ENDIF.

  ENDMETHOD.