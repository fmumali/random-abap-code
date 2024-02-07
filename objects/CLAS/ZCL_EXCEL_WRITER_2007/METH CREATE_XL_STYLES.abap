  METHOD create_xl_styles.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   dxfs-cellstyles are used in conditional formats:
*                CellIs, Expression, top10 ( forthcoming above average as well )
*                create own method to write dsfx-cellstyle to be reuseable by all these
*--------------------------------------------------------------------*


** Constant node name
    CONSTANTS: lc_xml_node_stylesheet        TYPE string VALUE 'styleSheet',
               " font
               lc_xml_node_fonts             TYPE string VALUE 'fonts',
               lc_xml_node_font              TYPE string VALUE 'font',
               lc_xml_node_color             TYPE string VALUE 'color',
               " fill
               lc_xml_node_fills             TYPE string VALUE 'fills',
               lc_xml_node_fill              TYPE string VALUE 'fill',
               lc_xml_node_patternfill       TYPE string VALUE 'patternFill',
               lc_xml_node_fgcolor           TYPE string VALUE 'fgColor',
               lc_xml_node_bgcolor           TYPE string VALUE 'bgColor',
               lc_xml_node_gradientfill      TYPE string VALUE 'gradientFill',
               lc_xml_node_stop              TYPE string VALUE 'stop',
               " borders
               lc_xml_node_borders           TYPE string VALUE 'borders',
               lc_xml_node_border            TYPE string VALUE 'border',
               lc_xml_node_left              TYPE string VALUE 'left',
               lc_xml_node_right             TYPE string VALUE 'right',
               lc_xml_node_top               TYPE string VALUE 'top',
               lc_xml_node_bottom            TYPE string VALUE 'bottom',
               lc_xml_node_diagonal          TYPE string VALUE 'diagonal',
               " numfmt
               lc_xml_node_numfmts           TYPE string VALUE 'numFmts',
               lc_xml_node_numfmt            TYPE string VALUE 'numFmt',
               " Styles
               lc_xml_node_cellstylexfs      TYPE string VALUE 'cellStyleXfs',
               lc_xml_node_xf                TYPE string VALUE 'xf',
               lc_xml_node_cellxfs           TYPE string VALUE 'cellXfs',
               lc_xml_node_cellstyles        TYPE string VALUE 'cellStyles',
               lc_xml_node_cellstyle         TYPE string VALUE 'cellStyle',
               lc_xml_node_dxfs              TYPE string VALUE 'dxfs',
               lc_xml_node_tablestyles       TYPE string VALUE 'tableStyles',
               " Colors
               lc_xml_node_colors            TYPE string VALUE 'colors',
               lc_xml_node_indexedcolors     TYPE string VALUE 'indexedColors',
               lc_xml_node_rgbcolor          TYPE string VALUE 'rgbColor',
               lc_xml_node_mrucolors         TYPE string VALUE 'mruColors',
               " Alignment
               lc_xml_node_alignment         TYPE string VALUE 'alignment',
               " Protection
               lc_xml_node_protection        TYPE string VALUE 'protection',
               " Node attributes
               lc_xml_attr_count             TYPE string VALUE 'count',
               lc_xml_attr_val               TYPE string VALUE 'val',
               lc_xml_attr_theme             TYPE string VALUE 'theme',
               lc_xml_attr_rgb               TYPE string VALUE 'rgb',
               lc_xml_attr_indexed           TYPE string VALUE 'indexed',
               lc_xml_attr_tint              TYPE string VALUE 'tint',
               lc_xml_attr_style             TYPE string VALUE 'style',
               lc_xml_attr_position          TYPE string VALUE 'position',
               lc_xml_attr_degree            TYPE string VALUE 'degree',
               lc_xml_attr_patterntype       TYPE string VALUE 'patternType',
               lc_xml_attr_numfmtid          TYPE string VALUE 'numFmtId',
               lc_xml_attr_fontid            TYPE string VALUE 'fontId',
               lc_xml_attr_fillid            TYPE string VALUE 'fillId',
               lc_xml_attr_borderid          TYPE string VALUE 'borderId',
               lc_xml_attr_xfid              TYPE string VALUE 'xfId',
               lc_xml_attr_applynumberformat TYPE string VALUE 'applyNumberFormat',
               lc_xml_attr_applyprotection   TYPE string VALUE 'applyProtection',
               lc_xml_attr_applyfont         TYPE string VALUE 'applyFont',
               lc_xml_attr_applyfill         TYPE string VALUE 'applyFill',
               lc_xml_attr_applyborder       TYPE string VALUE 'applyBorder',
               lc_xml_attr_name              TYPE string VALUE 'name',
               lc_xml_attr_builtinid         TYPE string VALUE 'builtinId',
               lc_xml_attr_defaulttablestyle TYPE string VALUE 'defaultTableStyle',
               lc_xml_attr_defaultpivotstyle TYPE string VALUE 'defaultPivotStyle',
               lc_xml_attr_applyalignment    TYPE string VALUE 'applyAlignment',
               lc_xml_attr_horizontal        TYPE string VALUE 'horizontal',
               lc_xml_attr_formatcode        TYPE string VALUE 'formatCode',
               lc_xml_attr_vertical          TYPE string VALUE 'vertical',
               lc_xml_attr_wraptext          TYPE string VALUE 'wrapText',
               lc_xml_attr_textrotation      TYPE string VALUE 'textRotation',
               lc_xml_attr_shrinktofit       TYPE string VALUE 'shrinkToFit',
               lc_xml_attr_indent            TYPE string VALUE 'indent',
               lc_xml_attr_locked            TYPE string VALUE 'locked',
               lc_xml_attr_hidden            TYPE string VALUE 'hidden',
               lc_xml_attr_diagonalup        TYPE string VALUE 'diagonalUp',
               lc_xml_attr_diagonaldown      TYPE string VALUE 'diagonalDown',
               " Node namespace
               lc_xml_node_ns                TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
               lc_xml_attr_type              TYPE string VALUE 'type',
               lc_xml_attr_bottom            TYPE string VALUE 'bottom',
               lc_xml_attr_top               TYPE string VALUE 'top',
               lc_xml_attr_right             TYPE string VALUE 'right',
               lc_xml_attr_left              TYPE string VALUE 'left'.

    DATA: lo_document        TYPE REF TO if_ixml_document,
          lo_element_root    TYPE REF TO if_ixml_element,
          lo_element_fonts   TYPE REF TO if_ixml_element,
          lo_element_font    TYPE REF TO if_ixml_element,
          lo_element_fills   TYPE REF TO if_ixml_element,
          lo_element_fill    TYPE REF TO if_ixml_element,
          lo_element_borders TYPE REF TO if_ixml_element,
          lo_element_border  TYPE REF TO if_ixml_element,
          lo_element_numfmts TYPE REF TO if_ixml_element,
          lo_element_numfmt  TYPE REF TO if_ixml_element,
          lo_element_cellxfs TYPE REF TO if_ixml_element,
          lo_element         TYPE REF TO if_ixml_element,
          lo_sub_element     TYPE REF TO if_ixml_element,
          lo_sub_element_2   TYPE REF TO if_ixml_element,
          lo_iterator        TYPE REF TO zcl_excel_collection_iterator,
          lo_iterator2       TYPE REF TO zcl_excel_collection_iterator,
          lo_worksheet       TYPE REF TO zcl_excel_worksheet,
          lo_style_cond      TYPE REF TO zcl_excel_style_cond,
          lo_style           TYPE REF TO zcl_excel_style.


    DATA: lt_fonts          TYPE zexcel_t_style_font,
          ls_font           TYPE zexcel_s_style_font,
          lt_fills          TYPE zexcel_t_style_fill,
          ls_fill           TYPE zexcel_s_style_fill,
          lt_borders        TYPE zexcel_t_style_border,
          ls_border         TYPE zexcel_s_style_border,
          lt_numfmts        TYPE zexcel_t_style_numfmt,
          ls_numfmt         TYPE zexcel_s_style_numfmt,
          lt_protections    TYPE zexcel_t_style_protection,
          ls_protection     TYPE zexcel_s_style_protection,
          lt_alignments     TYPE zexcel_t_style_alignment,
          ls_alignment      TYPE zexcel_s_style_alignment,
          lt_cellxfs        TYPE zexcel_t_cellxfs,
          ls_cellxfs        TYPE zexcel_s_cellxfs,
          ls_styles_mapping TYPE zexcel_s_styles_mapping,
          lt_colors         TYPE zexcel_t_style_color_argb,
          ls_color          LIKE LINE OF lt_colors.

    DATA: lv_value         TYPE string,
          lv_dfx_count     TYPE i,
          lv_fonts_count   TYPE i,
          lv_fills_count   TYPE i,
          lv_borders_count TYPE i,
          lv_cellxfs_count TYPE i.

    TYPES: BEGIN OF ts_built_in_format,
             num_format TYPE zexcel_number_format,
             id         TYPE i,
           END OF ts_built_in_format.

    DATA: lt_built_in_num_formats TYPE HASHED TABLE OF ts_built_in_format WITH UNIQUE KEY num_format,
          ls_built_in_num_format  LIKE LINE OF lt_built_in_num_formats.
    FIELD-SYMBOLS: <ls_built_in_format> LIKE LINE OF lt_built_in_num_formats,
                   <ls_reader_built_in> LIKE LINE OF zcl_excel_style_number_format=>mt_built_in_num_formats.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_stylesheet
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_ns ).

**********************************************************************
* STEP 4: Create subnodes

    lo_element_fonts = lo_document->create_simple_element( name   = lc_xml_node_fonts
                                                           parent = lo_document ).

    lo_element_fills = lo_document->create_simple_element( name   = lc_xml_node_fills
                                                           parent = lo_document ).

    lo_element_borders = lo_document->create_simple_element( name   = lc_xml_node_borders
                                                             parent = lo_document ).

    lo_element_cellxfs = lo_document->create_simple_element( name   = lc_xml_node_cellxfs
                                                             parent = lo_document ).

    lo_element_numfmts = lo_document->create_simple_element( name   = lc_xml_node_numfmts
                                                             parent = lo_document ).

* Prepare built-in number formats.
    LOOP AT zcl_excel_style_number_format=>mt_built_in_num_formats ASSIGNING <ls_reader_built_in>.
      ls_built_in_num_format-id         = <ls_reader_built_in>-id.
      ls_built_in_num_format-num_format = <ls_reader_built_in>-format->format_code.
      INSERT ls_built_in_num_format INTO TABLE lt_built_in_num_formats.
    ENDLOOP.
* Compress styles
    lo_iterator = excel->get_styles_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_style ?= lo_iterator->get_next( ).
      ls_font       = lo_style->font->get_structure( ).
      ls_fill       = lo_style->fill->get_structure( ).
      ls_border     = lo_style->borders->get_structure( ).
      ls_alignment  = lo_style->alignment->get_structure( ).
      ls_protection = lo_style->protection->get_structure( ).
      ls_numfmt     = lo_style->number_format->get_structure( ).

      CLEAR ls_cellxfs.


* Compress fonts
      READ TABLE lt_fonts FROM ls_font TRANSPORTING NO FIELDS.
      IF sy-subrc EQ 0.
        ls_cellxfs-fontid = sy-tabix.
      ELSE.
        APPEND ls_font TO lt_fonts.
        DESCRIBE TABLE lt_fonts LINES ls_cellxfs-fontid.
      ENDIF.
      SUBTRACT 1 FROM ls_cellxfs-fontid.

* Compress alignment
      READ TABLE lt_alignments FROM ls_alignment TRANSPORTING NO FIELDS.
      IF sy-subrc EQ 0.
        ls_cellxfs-alignmentid = sy-tabix.
      ELSE.
        APPEND ls_alignment TO lt_alignments.
        DESCRIBE TABLE lt_alignments LINES ls_cellxfs-alignmentid.
      ENDIF.
      SUBTRACT 1 FROM ls_cellxfs-alignmentid.

* Compress fills
      READ TABLE lt_fills FROM ls_fill TRANSPORTING NO FIELDS.
      IF sy-subrc EQ 0.
        ls_cellxfs-fillid = sy-tabix.
      ELSE.
        APPEND ls_fill TO lt_fills.
        DESCRIBE TABLE lt_fills LINES ls_cellxfs-fillid.
      ENDIF.
      SUBTRACT 1 FROM ls_cellxfs-fillid.

* Compress borders
      READ TABLE lt_borders FROM ls_border TRANSPORTING NO FIELDS.
      IF sy-subrc EQ 0.
        ls_cellxfs-borderid = sy-tabix.
      ELSE.
        APPEND ls_border TO lt_borders.
        DESCRIBE TABLE lt_borders LINES ls_cellxfs-borderid.
      ENDIF.
      SUBTRACT 1 FROM ls_cellxfs-borderid.

* Compress protection
      IF ls_protection-locked EQ c_on AND ls_protection-hidden EQ c_off.
        ls_cellxfs-applyprotection    = 0.
      ELSE.
        READ TABLE lt_protections FROM ls_protection TRANSPORTING NO FIELDS.
        IF sy-subrc EQ 0.
          ls_cellxfs-protectionid = sy-tabix.
        ELSE.
          APPEND ls_protection TO lt_protections.
          DESCRIBE TABLE lt_protections LINES ls_cellxfs-protectionid.
        ENDIF.
        ls_cellxfs-applyprotection    = 1.
      ENDIF.
      SUBTRACT 1 FROM ls_cellxfs-protectionid.

* Compress number formats

      "-----------
      IF ls_numfmt-numfmt NE zcl_excel_style_number_format=>c_format_date_std." and ls_numfmt-NUMFMT ne 'STD_NDEC'. " ALE Changes on going
        "---
        IF ls_numfmt IS NOT INITIAL.
* issue  #389 - Problem with built-in format ( those are not being taken account of )
* There are some internal number formats built-in into EXCEL
* Use these instead of duplicating the entries here, since they seem to be language-dependant and adjust to user settings in excel
          READ TABLE lt_built_in_num_formats ASSIGNING <ls_built_in_format> WITH TABLE KEY num_format = ls_numfmt-numfmt.
          IF sy-subrc = 0.
            ls_cellxfs-numfmtid = <ls_built_in_format>-id.
          ELSE.
            READ TABLE lt_numfmts FROM ls_numfmt TRANSPORTING NO FIELDS.
            IF sy-subrc EQ 0.
              ls_cellxfs-numfmtid = sy-tabix.
            ELSE.
              APPEND ls_numfmt TO lt_numfmts.
              DESCRIBE TABLE lt_numfmts LINES ls_cellxfs-numfmtid.
            ENDIF.
            ADD zcl_excel_common=>c_excel_numfmt_offset TO ls_cellxfs-numfmtid. " Add OXML offset for custom styles
          ENDIF.
          ls_cellxfs-applynumberformat    = 1.
        ELSE.
          ls_cellxfs-applynumberformat    = 0.
        ENDIF.
        "----------- " ALE changes on going
      ELSE.
        ls_cellxfs-applynumberformat    = 1.
        IF ls_numfmt-numfmt EQ zcl_excel_style_number_format=>c_format_date_std.
          ls_cellxfs-numfmtid = 14.
        ENDIF.
      ENDIF.
      "---

      IF ls_cellxfs-fontid NE 0.
        ls_cellxfs-applyfont    = 1.
      ELSE.
        ls_cellxfs-applyfont    = 0.
      ENDIF.
      IF ls_cellxfs-alignmentid NE 0.
        ls_cellxfs-applyalignment = 1.
      ELSE.
        ls_cellxfs-applyalignment = 0.
      ENDIF.
      IF ls_cellxfs-fillid NE 0.
        ls_cellxfs-applyfill    = 1.
      ELSE.
        ls_cellxfs-applyfill    = 0.
      ENDIF.
      IF ls_cellxfs-borderid NE 0.
        ls_cellxfs-applyborder    = 1.
      ELSE.
        ls_cellxfs-applyborder    = 0.
      ENDIF.

* Remap styles
      READ TABLE lt_cellxfs FROM ls_cellxfs TRANSPORTING NO FIELDS.
      IF sy-subrc EQ 0.
        ls_styles_mapping-style = sy-tabix.
      ELSE.
        APPEND ls_cellxfs TO lt_cellxfs.
        DESCRIBE TABLE lt_cellxfs LINES ls_styles_mapping-style.
      ENDIF.
      SUBTRACT 1 FROM ls_styles_mapping-style.
      ls_styles_mapping-guid = lo_style->get_guid( ).
      APPEND ls_styles_mapping TO me->styles_mapping.
    ENDWHILE.

    " create numfmt elements
    LOOP AT lt_numfmts INTO ls_numfmt.
      lo_element_numfmt = lo_document->create_simple_element( name   = lc_xml_node_numfmt
                                                              parent = lo_document ).
      lv_value = sy-tabix + zcl_excel_common=>c_excel_numfmt_offset.
      CONDENSE lv_value.
      lo_element_numfmt->set_attribute_ns( name  = lc_xml_attr_numfmtid
                                        value = lv_value ).
      lv_value = ls_numfmt-numfmt.
      lo_element_numfmt->set_attribute_ns( name  = lc_xml_attr_formatcode
                                           value = lv_value ).
      lo_element_numfmts->append_child( new_child = lo_element_numfmt ).
    ENDLOOP.

    " create font elements
    LOOP AT lt_fonts INTO ls_font.
      lo_element_font = lo_document->create_simple_element( name   = lc_xml_node_font
                                                            parent = lo_document ).
      create_xl_styles_font_node( io_document = lo_document
                                  io_parent   = lo_element_font
                                  is_font     = ls_font ).
      lo_element_fonts->append_child( new_child = lo_element_font ).
    ENDLOOP.

    " create fill elements
    LOOP AT lt_fills INTO ls_fill.
      lo_element_fill = lo_document->create_simple_element( name   = lc_xml_node_fill
                                                            parent = lo_document ).

      IF ls_fill-gradtype IS NOT INITIAL.
        "gradient

        lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_gradientfill
                                                            parent = lo_document ).
        IF ls_fill-gradtype-degree IS NOT INITIAL.
          lv_value = ls_fill-gradtype-degree.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_degree  value = lv_value ).
        ENDIF.
        IF ls_fill-gradtype-type IS NOT INITIAL.
          lv_value = ls_fill-gradtype-type.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_type  value = lv_value ).
        ENDIF.
        IF ls_fill-gradtype-bottom IS NOT INITIAL.
          lv_value = ls_fill-gradtype-bottom.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_bottom  value = lv_value ).
        ENDIF.
        IF ls_fill-gradtype-top IS NOT INITIAL.
          lv_value = ls_fill-gradtype-top.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_top  value = lv_value ).
        ENDIF.
        IF ls_fill-gradtype-right IS NOT INITIAL.
          lv_value = ls_fill-gradtype-right.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_right  value = lv_value ).
        ENDIF.
        IF ls_fill-gradtype-left IS NOT INITIAL.
          lv_value = ls_fill-gradtype-left.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_left  value = lv_value ).
        ENDIF.

        IF ls_fill-gradtype-position3 IS NOT INITIAL.
          "create <stop> elements for gradients, we can have 2 or 3 stops in each gradient
          lo_sub_element_2 =  lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                  parent = lo_sub_element ).
          lv_value = ls_fill-gradtype-position1.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).

          lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                 parent = lo_sub_element ).

          lv_value = ls_fill-gradtype-position2.

          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position
                                              value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-fgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).

          lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                 parent = lo_sub_element ).

          lv_value = ls_fill-gradtype-position3.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position
                                              value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).

        ELSE.
          "create <stop> elements for gradients, we can have 2 or 3 stops in each gradient
          lo_sub_element_2 =  lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                  parent = lo_sub_element ).
          lv_value = ls_fill-gradtype-position1.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).

          lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                 parent = lo_sub_element ).

          lv_value = ls_fill-gradtype-position2.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position
                                              value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-fgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).
        ENDIF.

      ELSE.
        "pattern
        lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_patternfill
                                                             parent = lo_document ).
        lv_value = ls_fill-filltype.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_patterntype
                                          value = lv_value ).
        " fgcolor
        create_xl_styles_color_node(
            io_document        = lo_document
            io_parent          = lo_sub_element
            is_color           = ls_fill-fgcolor
            iv_color_elem_name = lc_xml_node_fgcolor ).

        IF  ls_fill-fgcolor-rgb IS INITIAL AND
            ls_fill-fgcolor-indexed EQ zcl_excel_style_color=>c_indexed_not_set AND
            ls_fill-fgcolor-theme EQ zcl_excel_style_color=>c_theme_not_set AND
            ls_fill-fgcolor-tint IS INITIAL AND ls_fill-bgcolor-indexed EQ zcl_excel_style_color=>c_indexed_sys_foreground.

          " bgcolor
          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_bgcolor ).

        ENDIF.
      ENDIF.

      lo_element_fill->append_child( new_child = lo_sub_element )."pattern
      lo_element_fills->append_child( new_child = lo_element_fill ).
    ENDLOOP.

    " create border elements
    LOOP AT lt_borders INTO ls_border.
      lo_element_border = lo_document->create_simple_element( name   = lc_xml_node_border
                                                              parent = lo_document ).

      IF ls_border-diagonalup IS NOT INITIAL.
        lv_value = ls_border-diagonalup.
        CONDENSE lv_value.
        lo_element_border->set_attribute_ns( name  = lc_xml_attr_diagonalup
                                          value = lv_value ).
      ENDIF.

      IF ls_border-diagonaldown IS NOT INITIAL.
        lv_value = ls_border-diagonaldown.
        CONDENSE lv_value.
        lo_element_border->set_attribute_ns( name  = lc_xml_attr_diagonaldown
                                          value = lv_value ).
      ENDIF.

      "left
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_left
                                                           parent = lo_document ).
      IF ls_border-left_style IS NOT INITIAL.
        lv_value = ls_border-left_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      ENDIF.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-left_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).

      "right
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_right
                                                           parent = lo_document ).
      IF ls_border-right_style IS NOT INITIAL.
        lv_value = ls_border-right_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      ENDIF.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-right_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).

      "top
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_top
                                                           parent = lo_document ).
      IF ls_border-top_style IS NOT INITIAL.
        lv_value = ls_border-top_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      ENDIF.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-top_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).

      "bottom
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_bottom
                                                           parent = lo_document ).
      IF ls_border-bottom_style IS NOT INITIAL.
        lv_value = ls_border-bottom_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      ENDIF.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-bottom_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).

      "diagonal
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_diagonal
                                                           parent = lo_document ).
      IF ls_border-diagonal_style IS NOT INITIAL.
        lv_value = ls_border-diagonal_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      ENDIF.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-diagonal_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).
      lo_element_borders->append_child( new_child = lo_element_border ).
    ENDLOOP.

    " update attribute "count"
    DESCRIBE TABLE lt_fonts LINES lv_fonts_count.
    lv_value = lv_fonts_count.
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_element_fonts->set_attribute_ns( name  = lc_xml_attr_count
                                        value = lv_value ).
    DESCRIBE TABLE lt_fills LINES lv_fills_count.
    lv_value = lv_fills_count.
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_element_fills->set_attribute_ns( name  = lc_xml_attr_count
                                        value = lv_value ).
    DESCRIBE TABLE lt_borders LINES lv_borders_count.
    lv_value = lv_borders_count.
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_element_borders->set_attribute_ns( name  = lc_xml_attr_count
                                          value = lv_value ).
    DESCRIBE TABLE lt_cellxfs LINES lv_cellxfs_count.
    lv_value = lv_cellxfs_count.
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_element_cellxfs->set_attribute_ns( name  = lc_xml_attr_count
                                          value = lv_value ).

    " Append to root node
    lo_element_root->append_child( new_child = lo_element_numfmts ).
    lo_element_root->append_child( new_child = lo_element_fonts ).
    lo_element_root->append_child( new_child = lo_element_fills ).
    lo_element_root->append_child( new_child = lo_element_borders ).

    " cellstylexfs node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_cellstylexfs
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_count
                                  value = '1' ).
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_xf
                                                         parent = lo_document ).

    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_numfmtid
                                      value = c_off ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_fontid
                                      value = c_off ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_fillid
                                      value = c_off ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_borderid
                                      value = c_off ).

    lo_element->append_child( new_child = lo_sub_element ).
    lo_element_root->append_child( new_child = lo_element ).

    LOOP AT lt_cellxfs INTO ls_cellxfs.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_xf
                                                          parent = lo_document ).
      lv_value = ls_cellxfs-numfmtid.
      SHIFT lv_value RIGHT DELETING TRAILING space.
      SHIFT lv_value LEFT DELETING LEADING space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_numfmtid
                                    value = lv_value ).
      lv_value = ls_cellxfs-fontid.
      SHIFT lv_value RIGHT DELETING TRAILING space.
      SHIFT lv_value LEFT DELETING LEADING space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_fontid
                                    value = lv_value ).
      lv_value = ls_cellxfs-fillid.
      SHIFT lv_value RIGHT DELETING TRAILING space.
      SHIFT lv_value LEFT DELETING LEADING space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_fillid
                                    value = lv_value ).
      lv_value = ls_cellxfs-borderid.
      SHIFT lv_value RIGHT DELETING TRAILING space.
      SHIFT lv_value LEFT DELETING LEADING space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_borderid
                                    value = lv_value ).
      lv_value = ls_cellxfs-xfid.
      SHIFT lv_value RIGHT DELETING TRAILING space.
      SHIFT lv_value LEFT DELETING LEADING space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_xfid
                                    value = lv_value ).
      IF ls_cellxfs-applynumberformat EQ 1.
        lv_value = ls_cellxfs-applynumberformat.
        SHIFT lv_value RIGHT DELETING TRAILING space.
        SHIFT lv_value LEFT DELETING LEADING space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applynumberformat
                                      value = lv_value ).
      ENDIF.
      IF ls_cellxfs-applyfont EQ 1.
        lv_value = ls_cellxfs-applyfont.
        SHIFT lv_value RIGHT DELETING TRAILING space.
        SHIFT lv_value LEFT DELETING LEADING space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyfont
                                      value = lv_value ).
      ENDIF.
      IF ls_cellxfs-applyfill EQ 1.
        lv_value = ls_cellxfs-applyfill.
        SHIFT lv_value RIGHT DELETING TRAILING space.
        SHIFT lv_value LEFT DELETING LEADING space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyfill
                                      value = lv_value ).
      ENDIF.
      IF ls_cellxfs-applyborder EQ 1.
        lv_value = ls_cellxfs-applyborder.
        SHIFT lv_value RIGHT DELETING TRAILING space.
        SHIFT lv_value LEFT DELETING LEADING space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyborder
                                      value = lv_value ).
      ENDIF.
      IF ls_cellxfs-applyalignment EQ 1. " depends on each style not for all the sheet
        lv_value = ls_cellxfs-applyalignment.
        SHIFT lv_value RIGHT DELETING TRAILING space.
        SHIFT lv_value LEFT DELETING LEADING space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyalignment
                                      value = lv_value ).
        lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_alignment
                                                               parent = lo_document ).
        ADD 1 TO ls_cellxfs-alignmentid. "Table index starts from 1
        READ TABLE lt_alignments INTO ls_alignment INDEX ls_cellxfs-alignmentid.
        SUBTRACT 1 FROM ls_cellxfs-alignmentid.
        IF ls_alignment-horizontal IS NOT INITIAL.
          lv_value = ls_alignment-horizontal.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_horizontal
                                              value = lv_value ).
        ENDIF.
        IF ls_alignment-vertical IS NOT INITIAL.
          lv_value = ls_alignment-vertical.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_vertical
                                              value = lv_value ).
        ENDIF.
        IF ls_alignment-wraptext EQ abap_true.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_wraptext
                                              value = c_on ).
        ENDIF.
        IF ls_alignment-textrotation IS NOT INITIAL.
          lv_value = ls_alignment-textrotation.
          SHIFT lv_value RIGHT DELETING TRAILING space.
          SHIFT lv_value LEFT DELETING LEADING space.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_textrotation
                                              value = lv_value ).
        ENDIF.
        IF ls_alignment-shrinktofit EQ abap_true.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_shrinktofit
                                              value = c_on ).
        ENDIF.
        IF ls_alignment-indent IS NOT INITIAL.
          lv_value = ls_alignment-indent.
          SHIFT lv_value RIGHT DELETING TRAILING space.
          SHIFT lv_value LEFT DELETING LEADING space.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_indent
                                              value = lv_value ).
        ENDIF.

        lo_element->append_child( new_child = lo_sub_element_2 ).
      ENDIF.
      IF ls_cellxfs-applyprotection EQ 1.
        lv_value = ls_cellxfs-applyprotection.
        CONDENSE lv_value NO-GAPS.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyprotection
                                      value = lv_value ).
        lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_protection
                                                               parent = lo_document ).
        ADD 1 TO ls_cellxfs-protectionid. "Table index starts from 1
        READ TABLE lt_protections INTO ls_protection INDEX ls_cellxfs-protectionid.
        SUBTRACT 1 FROM ls_cellxfs-protectionid.
        IF ls_protection-locked IS NOT INITIAL.
          lv_value = ls_protection-locked.
          CONDENSE lv_value.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_locked
                                              value = lv_value ).
        ENDIF.
        IF ls_protection-hidden IS NOT INITIAL.
          lv_value = ls_protection-hidden.
          CONDENSE lv_value.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_hidden
                                              value = lv_value ).
        ENDIF.
        lo_element->append_child( new_child = lo_sub_element_2 ).
      ENDIF.
      lo_element_cellxfs->append_child( new_child = lo_element ).
    ENDLOOP.

    lo_element_root->append_child( new_child = lo_element_cellxfs ).

    " cellStyles node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_cellstyles
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_count
                                  value = '1' ).
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_cellstyle
                                                         parent = lo_document ).

    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                      value = 'Normal' ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_xfid
                                      value = c_off ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_builtinid
                                      value = c_off ).

    lo_element->append_child( new_child = lo_sub_element ).
    lo_element_root->append_child( new_child = lo_element ).

    " dxfs node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_dxfs
                                                     parent = lo_document ).

    lo_iterator = me->excel->get_worksheets_iterator( ).
    " get sheets
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).
      " Conditional formatting styles into exch sheet
      lo_iterator2 = lo_worksheet->get_style_cond_iterator( ).
      WHILE lo_iterator2->has_next( ) EQ abap_true.
        lo_style_cond ?= lo_iterator2->get_next( ).
        CASE lo_style_cond->rule.
* begin of change issue #366 - missing conditional rules: top10, move dfx-styles to own method
          WHEN zcl_excel_style_cond=>c_rule_cellis.
            me->create_dxf_style( EXPORTING
                                    iv_cell_style    = lo_style_cond->mode_cellis-cell_style
                                    io_dxf_element   = lo_element
                                    io_ixml_document = lo_document
                                    it_cellxfs       = lt_cellxfs
                                    it_fonts         = lt_fonts
                                    it_fills         = lt_fills
                                  CHANGING
                                    cv_dfx_count     = lv_dfx_count ).

          WHEN zcl_excel_style_cond=>c_rule_expression.
            me->create_dxf_style( EXPORTING
                          iv_cell_style    = lo_style_cond->mode_expression-cell_style
                          io_dxf_element   = lo_element
                          io_ixml_document = lo_document
                          it_cellxfs       = lt_cellxfs
                          it_fonts         = lt_fonts
                          it_fills         = lt_fills
                        CHANGING
                          cv_dfx_count     = lv_dfx_count ).



          WHEN zcl_excel_style_cond=>c_rule_top10.
            me->create_dxf_style( EXPORTING
                                    iv_cell_style    = lo_style_cond->mode_top10-cell_style
                                    io_dxf_element   = lo_element
                                    io_ixml_document = lo_document
                                    it_cellxfs       = lt_cellxfs
                                    it_fonts         = lt_fonts
                                    it_fills         = lt_fills
                                  CHANGING
                                    cv_dfx_count     = lv_dfx_count ).

          WHEN zcl_excel_style_cond=>c_rule_above_average.
            me->create_dxf_style( EXPORTING
                                    iv_cell_style    = lo_style_cond->mode_above_average-cell_style
                                    io_dxf_element   = lo_element
                                    io_ixml_document = lo_document
                                    it_cellxfs       = lt_cellxfs
                                    it_fonts         = lt_fonts
                                    it_fills         = lt_fills
                                  CHANGING
                                    cv_dfx_count     = lv_dfx_count ).
* begin of change issue #366 - missing conditional rules: top10, move dfx-styles to own method

          WHEN OTHERS.
            CONTINUE.
        ENDCASE.
      ENDWHILE.
    ENDWHILE.

    lv_value = lv_dfx_count.
    CONDENSE lv_value.
    lo_element->set_attribute_ns( name  = lc_xml_attr_count
                                  value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " tableStyles node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_tablestyles
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_count
                                  value = '0' ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_defaulttablestyle
                                  value = zcl_excel_table=>builtinstyle_medium9 ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_defaultpivotstyle
                                  value = zcl_excel_table=>builtinstyle_pivot_light16 ).
    lo_element_root->append_child( new_child = lo_element ).

    "write legacy color palette in case any indexed color was changed
    IF excel->legacy_palette->is_modified( ) = abap_true.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_colors
                                                     parent   = lo_document ).
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_indexedcolors
                                                         parent   = lo_document ).
      lo_element->append_child( new_child = lo_sub_element ).

      lt_colors = excel->legacy_palette->get_colors( ).
      LOOP AT lt_colors INTO ls_color.
        lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_rgbcolor
                                                               parent = lo_document ).
        lv_value = ls_color.
        lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_rgb
                                            value = lv_value ).
        lo_sub_element->append_child( new_child = lo_sub_element_2 ).
      ENDLOOP.

      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.