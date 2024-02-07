  METHOD zif_excel_style_changer~apply.

    DATA: stylemapping TYPE zexcel_s_stylemapping,
          lo_worksheet TYPE REF TO zcl_excel_worksheet,
          l_guid       TYPE zexcel_cell_style.

    lo_worksheet = excel->get_worksheet_by_name( ip_sheet_name = ip_worksheet->get_title( ) ).
    IF lo_worksheet <> ip_worksheet.
      zcx_excel=>raise_text( 'Worksheet doesn''t correspond to workbook of style changer'(001) ).
    ENDIF.

    TRY.
        ip_worksheet->get_cell( EXPORTING ip_column = ip_column
                                          ip_row    = ip_row
                                IMPORTING ep_guid   = l_guid ).
        stylemapping = excel->get_style_to_guid( l_guid ).
      CATCH zcx_excel.
* Error --> use submitted style
    ENDTRY.


    IF multiple_change_requested-complete = abap_true.
      stylemapping-complete_style = complete_style.
      stylemapping-complete_stylex = complete_stylex.
    ENDIF.

    IF multiple_change_requested-font = abap_true.
      stylemapping-complete_style-font = complete_style-font.
      stylemapping-complete_stylex-font = complete_stylex-font.
    ENDIF.

    IF multiple_change_requested-fill = abap_true.
      stylemapping-complete_style-fill = complete_style-fill.
      stylemapping-complete_stylex-fill = complete_stylex-fill.
    ENDIF.

    IF multiple_change_requested-borders-complete = abap_true.
      stylemapping-complete_style-borders = complete_style-borders.
      stylemapping-complete_stylex-borders = complete_stylex-borders.
    ENDIF.

    IF multiple_change_requested-borders-allborders = abap_true.
      stylemapping-complete_style-borders-allborders = complete_style-borders-allborders.
      stylemapping-complete_stylex-borders-allborders = complete_stylex-borders-allborders.
    ENDIF.

    IF multiple_change_requested-borders-diagonal = abap_true.
      stylemapping-complete_style-borders-diagonal = complete_style-borders-diagonal.
      stylemapping-complete_stylex-borders-diagonal = complete_stylex-borders-diagonal.
    ENDIF.

    IF multiple_change_requested-borders-down = abap_true.
      stylemapping-complete_style-borders-down = complete_style-borders-down.
      stylemapping-complete_stylex-borders-down = complete_stylex-borders-down.
    ENDIF.

    IF multiple_change_requested-borders-left = abap_true.
      stylemapping-complete_style-borders-left = complete_style-borders-left.
      stylemapping-complete_stylex-borders-left = complete_stylex-borders-left.
    ENDIF.

    IF multiple_change_requested-borders-right = abap_true.
      stylemapping-complete_style-borders-right = complete_style-borders-right.
      stylemapping-complete_stylex-borders-right = complete_stylex-borders-right.
    ENDIF.

    IF multiple_change_requested-borders-top = abap_true.
      stylemapping-complete_style-borders-top = complete_style-borders-top.
      stylemapping-complete_stylex-borders-top = complete_stylex-borders-top.
    ENDIF.

    IF multiple_change_requested-alignment = abap_true.
      stylemapping-complete_style-alignment = complete_style-alignment.
      stylemapping-complete_stylex-alignment = complete_stylex-alignment.
    ENDIF.

    IF multiple_change_requested-protection = abap_true.
      stylemapping-complete_style-protection = complete_style-protection.
      stylemapping-complete_stylex-protection = complete_stylex-protection.
    ENDIF.

    IF complete_stylex-number_format = abap_true.
      stylemapping-complete_style-number_format-format_code = complete_style-number_format-format_code.
      stylemapping-complete_stylex-number_format-format_code = abap_true.
    ENDIF.
    IF complete_stylex-font-bold = abap_true.
      stylemapping-complete_style-font-bold = complete_style-font-bold.
      stylemapping-complete_stylex-font-bold = complete_stylex-font-bold.
    ENDIF.
    IF complete_stylex-font-color = abap_true.
      stylemapping-complete_style-font-color = complete_style-font-color.
      stylemapping-complete_stylex-font-color = complete_stylex-font-color.
    ENDIF.
    IF complete_stylex-font-color-rgb = abap_true.
      stylemapping-complete_style-font-color-rgb = complete_style-font-color-rgb.
      stylemapping-complete_stylex-font-color-rgb = complete_stylex-font-color-rgb.
    ENDIF.
    IF complete_stylex-font-color-indexed = abap_true.
      stylemapping-complete_style-font-color-indexed = complete_style-font-color-indexed.
      stylemapping-complete_stylex-font-color-indexed = complete_stylex-font-color-indexed.
    ENDIF.
    IF complete_stylex-font-color-theme = abap_true.
      stylemapping-complete_style-font-color-theme = complete_style-font-color-theme.
      stylemapping-complete_stylex-font-color-theme = complete_stylex-font-color-theme.
    ENDIF.
    IF complete_stylex-font-color-tint = abap_true.
      stylemapping-complete_style-font-color-tint = complete_style-font-color-tint.
      stylemapping-complete_stylex-font-color-tint = complete_stylex-font-color-tint.
    ENDIF.
    IF complete_stylex-font-family = abap_true.
      stylemapping-complete_style-font-family = complete_style-font-family.
      stylemapping-complete_stylex-font-family = complete_stylex-font-family.
    ENDIF.
    IF complete_stylex-font-italic = abap_true.
      stylemapping-complete_style-font-italic = complete_style-font-italic.
      stylemapping-complete_stylex-font-italic = complete_stylex-font-italic.
    ENDIF.
    IF complete_stylex-font-name = abap_true.
      stylemapping-complete_style-font-name = complete_style-font-name.
      stylemapping-complete_stylex-font-name = complete_stylex-font-name.
    ENDIF.
    IF complete_stylex-font-scheme = abap_true.
      stylemapping-complete_style-font-scheme = complete_style-font-scheme.
      stylemapping-complete_stylex-font-scheme = complete_stylex-font-scheme.
    ENDIF.
    IF complete_stylex-font-size = abap_true.
      stylemapping-complete_style-font-size = complete_style-font-size.
      stylemapping-complete_stylex-font-size = complete_stylex-font-size.
    ENDIF.
    IF complete_stylex-font-strikethrough = abap_true.
      stylemapping-complete_style-font-strikethrough = complete_style-font-strikethrough.
      stylemapping-complete_stylex-font-strikethrough = complete_stylex-font-strikethrough.
    ENDIF.
    IF complete_stylex-font-underline = abap_true.
      stylemapping-complete_style-font-underline = complete_style-font-underline.
      stylemapping-complete_stylex-font-underline = complete_stylex-font-underline.
    ENDIF.
    IF complete_stylex-font-underline_mode = abap_true.
      stylemapping-complete_style-font-underline_mode = complete_style-font-underline_mode.
      stylemapping-complete_stylex-font-underline_mode = complete_stylex-font-underline_mode.
    ENDIF.

    IF complete_stylex-fill-filltype = abap_true.
      stylemapping-complete_style-fill-filltype = complete_style-fill-filltype.
      stylemapping-complete_stylex-fill-filltype = complete_stylex-fill-filltype.
    ENDIF.
    IF complete_stylex-fill-rotation = abap_true.
      stylemapping-complete_style-fill-rotation = complete_style-fill-rotation.
      stylemapping-complete_stylex-fill-rotation = complete_stylex-fill-rotation.
    ENDIF.
    IF complete_stylex-fill-fgcolor = abap_true.
      stylemapping-complete_style-fill-fgcolor = complete_style-fill-fgcolor.
      stylemapping-complete_stylex-fill-fgcolor = complete_stylex-fill-fgcolor.
    ENDIF.
    IF complete_stylex-fill-fgcolor-rgb = abap_true.
      stylemapping-complete_style-fill-fgcolor-rgb = complete_style-fill-fgcolor-rgb.
      stylemapping-complete_stylex-fill-fgcolor-rgb = complete_stylex-fill-fgcolor-rgb.
    ENDIF.
    IF complete_stylex-fill-fgcolor-indexed = abap_true.
      stylemapping-complete_style-fill-fgcolor-indexed = complete_style-fill-fgcolor-indexed.
      stylemapping-complete_stylex-fill-fgcolor-indexed = complete_stylex-fill-fgcolor-indexed.
    ENDIF.
    IF complete_stylex-fill-fgcolor-theme = abap_true.
      stylemapping-complete_style-fill-fgcolor-theme = complete_style-fill-fgcolor-theme.
      stylemapping-complete_stylex-fill-fgcolor-theme = complete_stylex-fill-fgcolor-theme.
    ENDIF.
    IF complete_stylex-fill-fgcolor-tint = abap_true.
      stylemapping-complete_style-fill-fgcolor-tint = complete_style-fill-fgcolor-tint.
      stylemapping-complete_stylex-fill-fgcolor-tint = complete_stylex-fill-fgcolor-tint.
    ENDIF.

    IF complete_stylex-fill-bgcolor = abap_true.
      stylemapping-complete_style-fill-bgcolor = complete_style-fill-bgcolor.
      stylemapping-complete_stylex-fill-bgcolor = complete_stylex-fill-bgcolor.
    ENDIF.
    IF complete_stylex-fill-bgcolor-rgb = abap_true.
      stylemapping-complete_style-fill-bgcolor-rgb = complete_style-fill-bgcolor-rgb.
      stylemapping-complete_stylex-fill-bgcolor-rgb = complete_stylex-fill-bgcolor-rgb.
    ENDIF.
    IF complete_stylex-fill-bgcolor-indexed = abap_true.
      stylemapping-complete_style-fill-bgcolor-indexed = complete_style-fill-bgcolor-indexed.
      stylemapping-complete_stylex-fill-bgcolor-indexed = complete_stylex-fill-bgcolor-indexed.
    ENDIF.
    IF complete_stylex-fill-bgcolor-theme = abap_true.
      stylemapping-complete_style-fill-bgcolor-theme = complete_style-fill-bgcolor-theme.
      stylemapping-complete_stylex-fill-bgcolor-theme = complete_stylex-fill-bgcolor-theme.
    ENDIF.
    IF complete_stylex-fill-bgcolor-tint = abap_true.
      stylemapping-complete_style-fill-bgcolor-tint = complete_style-fill-bgcolor-tint.
      stylemapping-complete_stylex-fill-bgcolor-tint = complete_stylex-fill-bgcolor-tint.
    ENDIF.

    IF complete_stylex-fill-gradtype-type = abap_true.
      stylemapping-complete_style-fill-gradtype-type = complete_style-fill-gradtype-type.
      stylemapping-complete_stylex-fill-gradtype-type = complete_stylex-fill-gradtype-type.
    ENDIF.
    IF complete_stylex-fill-gradtype-degree = abap_true.
      stylemapping-complete_style-fill-gradtype-degree = complete_style-fill-gradtype-degree.
      stylemapping-complete_stylex-fill-gradtype-degree = complete_stylex-fill-gradtype-degree.
    ENDIF.
    IF complete_stylex-fill-gradtype-bottom = abap_true.
      stylemapping-complete_style-fill-gradtype-bottom = complete_style-fill-gradtype-bottom.
      stylemapping-complete_stylex-fill-gradtype-bottom = complete_stylex-fill-gradtype-bottom.
    ENDIF.
    IF complete_stylex-fill-gradtype-left = abap_true.
      stylemapping-complete_style-fill-gradtype-left = complete_style-fill-gradtype-left.
      stylemapping-complete_stylex-fill-gradtype-left = complete_stylex-fill-gradtype-left.
    ENDIF.
    IF complete_stylex-fill-gradtype-top = abap_true.
      stylemapping-complete_style-fill-gradtype-top = complete_style-fill-gradtype-top.
      stylemapping-complete_stylex-fill-gradtype-top = complete_stylex-fill-gradtype-top.
    ENDIF.
    IF complete_stylex-fill-gradtype-right = abap_true.
      stylemapping-complete_style-fill-gradtype-right = complete_style-fill-gradtype-right.
      stylemapping-complete_stylex-fill-gradtype-right = complete_stylex-fill-gradtype-right.
    ENDIF.
    IF complete_stylex-fill-gradtype-position1 = abap_true.
      stylemapping-complete_style-fill-gradtype-position1 = complete_style-fill-gradtype-position1.
      stylemapping-complete_stylex-fill-gradtype-position1 = complete_stylex-fill-gradtype-position1.
    ENDIF.
    IF complete_stylex-fill-gradtype-position2 = abap_true.
      stylemapping-complete_style-fill-gradtype-position2 = complete_style-fill-gradtype-position2.
      stylemapping-complete_stylex-fill-gradtype-position2 = complete_stylex-fill-gradtype-position2.
    ENDIF.
    IF complete_stylex-fill-gradtype-position3 = abap_true.
      stylemapping-complete_style-fill-gradtype-position3 = complete_style-fill-gradtype-position3.
      stylemapping-complete_stylex-fill-gradtype-position3 = complete_stylex-fill-gradtype-position3.
    ENDIF.



    IF complete_stylex-borders-diagonal_mode = abap_true.
      stylemapping-complete_style-borders-diagonal_mode = complete_style-borders-diagonal_mode.
      stylemapping-complete_stylex-borders-diagonal_mode = complete_stylex-borders-diagonal_mode.
    ENDIF.
    IF complete_stylex-alignment-horizontal = abap_true.
      stylemapping-complete_style-alignment-horizontal = complete_style-alignment-horizontal.
      stylemapping-complete_stylex-alignment-horizontal = complete_stylex-alignment-horizontal.
    ENDIF.
    IF complete_stylex-alignment-vertical = abap_true.
      stylemapping-complete_style-alignment-vertical = complete_style-alignment-vertical.
      stylemapping-complete_stylex-alignment-vertical = complete_stylex-alignment-vertical.
    ENDIF.
    IF complete_stylex-alignment-textrotation = abap_true.
      stylemapping-complete_style-alignment-textrotation = complete_style-alignment-textrotation.
      stylemapping-complete_stylex-alignment-textrotation = complete_stylex-alignment-textrotation.
    ENDIF.
    IF complete_stylex-alignment-wraptext = abap_true.
      stylemapping-complete_style-alignment-wraptext = complete_style-alignment-wraptext.
      stylemapping-complete_stylex-alignment-wraptext = complete_stylex-alignment-wraptext.
    ENDIF.
    IF complete_stylex-alignment-shrinktofit = abap_true.
      stylemapping-complete_style-alignment-shrinktofit = complete_style-alignment-shrinktofit.
      stylemapping-complete_stylex-alignment-shrinktofit = complete_stylex-alignment-shrinktofit.
    ENDIF.
    IF complete_stylex-alignment-indent = abap_true.
      stylemapping-complete_style-alignment-indent = complete_style-alignment-indent.
      stylemapping-complete_stylex-alignment-indent = complete_stylex-alignment-indent.
    ENDIF.
    IF complete_stylex-protection-hidden = abap_true.
      stylemapping-complete_style-protection-hidden = complete_style-protection-hidden.
      stylemapping-complete_stylex-protection-hidden = complete_stylex-protection-hidden.
    ENDIF.
    IF complete_stylex-protection-locked = abap_true.
      stylemapping-complete_style-protection-locked = complete_style-protection-locked.
      stylemapping-complete_stylex-protection-locked = complete_stylex-protection-locked.
    ENDIF.

    IF complete_stylex-borders-allborders-border_style = abap_true.
      stylemapping-complete_style-borders-allborders-border_style = complete_style-borders-allborders-border_style.
      stylemapping-complete_stylex-borders-allborders-border_style = complete_stylex-borders-allborders-border_style.
    ENDIF.
    IF complete_stylex-borders-allborders-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-allborders-border_color-rgb = complete_style-borders-allborders-border_color-rgb.
      stylemapping-complete_stylex-borders-allborders-border_color-rgb = complete_stylex-borders-allborders-border_color-rgb.
    ENDIF.
    IF complete_stylex-borders-allborders-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-allborders-border_color-indexed = complete_style-borders-allborders-border_color-indexed.
      stylemapping-complete_stylex-borders-allborders-border_color-indexed = complete_stylex-borders-allborders-border_color-indexed.
    ENDIF.
    IF complete_stylex-borders-allborders-border_color-theme = abap_true.
      stylemapping-complete_style-borders-allborders-border_color-theme = complete_style-borders-allborders-border_color-theme.
      stylemapping-complete_stylex-borders-allborders-border_color-theme = complete_stylex-borders-allborders-border_color-theme.
    ENDIF.
    IF complete_stylex-borders-allborders-border_color-tint = abap_true.
      stylemapping-complete_style-borders-allborders-border_color-tint = complete_style-borders-allborders-border_color-tint.
      stylemapping-complete_stylex-borders-allborders-border_color-tint = complete_stylex-borders-allborders-border_color-tint.
    ENDIF.

    IF complete_stylex-borders-diagonal-border_style = abap_true.
      stylemapping-complete_style-borders-diagonal-border_style = complete_style-borders-diagonal-border_style.
      stylemapping-complete_stylex-borders-diagonal-border_style = complete_stylex-borders-diagonal-border_style.
    ENDIF.
    IF complete_stylex-borders-diagonal-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-diagonal-border_color-rgb = complete_style-borders-diagonal-border_color-rgb.
      stylemapping-complete_stylex-borders-diagonal-border_color-rgb = complete_stylex-borders-diagonal-border_color-rgb.
    ENDIF.
    IF complete_stylex-borders-diagonal-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-diagonal-border_color-indexed = complete_style-borders-diagonal-border_color-indexed.
      stylemapping-complete_stylex-borders-diagonal-border_color-indexed = complete_stylex-borders-diagonal-border_color-indexed.
    ENDIF.
    IF complete_stylex-borders-diagonal-border_color-theme = abap_true.
      stylemapping-complete_style-borders-diagonal-border_color-theme = complete_style-borders-diagonal-border_color-theme.
      stylemapping-complete_stylex-borders-diagonal-border_color-theme = complete_stylex-borders-diagonal-border_color-theme.
    ENDIF.
    IF complete_stylex-borders-diagonal-border_color-tint = abap_true.
      stylemapping-complete_style-borders-diagonal-border_color-tint = complete_style-borders-diagonal-border_color-tint.
      stylemapping-complete_stylex-borders-diagonal-border_color-tint = complete_stylex-borders-diagonal-border_color-tint.
    ENDIF.

    IF complete_stylex-borders-down-border_style = abap_true.
      stylemapping-complete_style-borders-down-border_style = complete_style-borders-down-border_style.
      stylemapping-complete_stylex-borders-down-border_style = complete_stylex-borders-down-border_style.
    ENDIF.
    IF complete_stylex-borders-down-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-down-border_color-rgb = complete_style-borders-down-border_color-rgb.
      stylemapping-complete_stylex-borders-down-border_color-rgb = complete_stylex-borders-down-border_color-rgb.
    ENDIF.
    IF complete_stylex-borders-down-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-down-border_color-indexed = complete_style-borders-down-border_color-indexed.
      stylemapping-complete_stylex-borders-down-border_color-indexed = complete_stylex-borders-down-border_color-indexed.
    ENDIF.
    IF complete_stylex-borders-down-border_color-theme = abap_true.
      stylemapping-complete_style-borders-down-border_color-theme = complete_style-borders-down-border_color-theme.
      stylemapping-complete_stylex-borders-down-border_color-theme = complete_stylex-borders-down-border_color-theme.
    ENDIF.
    IF complete_stylex-borders-down-border_color-tint = abap_true.
      stylemapping-complete_style-borders-down-border_color-tint = complete_style-borders-down-border_color-tint.
      stylemapping-complete_stylex-borders-down-border_color-tint = complete_stylex-borders-down-border_color-tint.
    ENDIF.

    IF complete_stylex-borders-left-border_style = abap_true.
      stylemapping-complete_style-borders-left-border_style = complete_style-borders-left-border_style.
      stylemapping-complete_stylex-borders-left-border_style = complete_stylex-borders-left-border_style.
    ENDIF.
    IF complete_stylex-borders-left-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-left-border_color-rgb = complete_style-borders-left-border_color-rgb.
      stylemapping-complete_stylex-borders-left-border_color-rgb = complete_stylex-borders-left-border_color-rgb.
    ENDIF.
    IF complete_stylex-borders-left-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-left-border_color-indexed = complete_style-borders-left-border_color-indexed.
      stylemapping-complete_stylex-borders-left-border_color-indexed = complete_stylex-borders-left-border_color-indexed.
    ENDIF.
    IF complete_stylex-borders-left-border_color-theme = abap_true.
      stylemapping-complete_style-borders-left-border_color-theme = complete_style-borders-left-border_color-theme.
      stylemapping-complete_stylex-borders-left-border_color-theme = complete_stylex-borders-left-border_color-theme.
    ENDIF.
    IF complete_stylex-borders-left-border_color-tint = abap_true.
      stylemapping-complete_style-borders-left-border_color-tint = complete_style-borders-left-border_color-tint.
      stylemapping-complete_stylex-borders-left-border_color-tint = complete_stylex-borders-left-border_color-tint.
    ENDIF.

    IF complete_stylex-borders-right-border_style = abap_true.
      stylemapping-complete_style-borders-right-border_style = complete_style-borders-right-border_style.
      stylemapping-complete_stylex-borders-right-border_style = complete_stylex-borders-right-border_style.
    ENDIF.
    IF complete_stylex-borders-right-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-right-border_color-rgb = complete_style-borders-right-border_color-rgb.
      stylemapping-complete_stylex-borders-right-border_color-rgb = complete_stylex-borders-right-border_color-rgb.
    ENDIF.
    IF complete_stylex-borders-right-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-right-border_color-indexed = complete_style-borders-right-border_color-indexed.
      stylemapping-complete_stylex-borders-right-border_color-indexed = complete_stylex-borders-right-border_color-indexed.
    ENDIF.
    IF complete_stylex-borders-right-border_color-theme = abap_true.
      stylemapping-complete_style-borders-right-border_color-theme = complete_style-borders-right-border_color-theme.
      stylemapping-complete_stylex-borders-right-border_color-theme = complete_stylex-borders-right-border_color-theme.
    ENDIF.
    IF complete_stylex-borders-right-border_color-tint = abap_true.
      stylemapping-complete_style-borders-right-border_color-tint = complete_style-borders-right-border_color-tint.
      stylemapping-complete_stylex-borders-right-border_color-tint = complete_stylex-borders-right-border_color-tint.
    ENDIF.

    IF complete_stylex-borders-top-border_style = abap_true.
      stylemapping-complete_style-borders-top-border_style = complete_style-borders-top-border_style.
      stylemapping-complete_stylex-borders-top-border_style = complete_stylex-borders-top-border_style.
    ENDIF.
    IF complete_stylex-borders-top-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-top-border_color-rgb = complete_style-borders-top-border_color-rgb.
      stylemapping-complete_stylex-borders-top-border_color-rgb = complete_stylex-borders-top-border_color-rgb.
    ENDIF.
    IF complete_stylex-borders-top-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-top-border_color-indexed = complete_style-borders-top-border_color-indexed.
      stylemapping-complete_stylex-borders-top-border_color-indexed = complete_stylex-borders-top-border_color-indexed.
    ENDIF.
    IF complete_stylex-borders-top-border_color-theme = abap_true.
      stylemapping-complete_style-borders-top-border_color-theme = complete_style-borders-top-border_color-theme.
      stylemapping-complete_stylex-borders-top-border_color-theme = complete_stylex-borders-top-border_color-theme.
    ENDIF.
    IF complete_stylex-borders-top-border_color-tint = abap_true.
      stylemapping-complete_style-borders-top-border_color-tint = complete_style-borders-top-border_color-tint.
      stylemapping-complete_stylex-borders-top-border_color-tint = complete_stylex-borders-top-border_color-tint.
    ENDIF.


* Now we have a completly filled styles.
* This can be used to get the guid
* Return guid if requested.  Might be used if copy&paste of styles is requested
    ep_guid = me->excel->get_static_cellstyle_guid( ip_cstyle_complete  = stylemapping-complete_style
                                                   ip_cstylex_complete = stylemapping-complete_stylex  ).
    lo_worksheet->set_cell_style( ip_column = ip_column
                                  ip_row    = ip_row
                                  ip_style  = ep_guid ).

  ENDMETHOD.