  METHOD load_styles.

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (wip )              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          aligning code
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
    TYPES: BEGIN OF lty_xf,
             applyalignment    TYPE string,
             applyborder       TYPE string,
             applyfill         TYPE string,
             applyfont         TYPE string,
             applynumberformat TYPE string,
             applyprotection   TYPE string,
             borderid          TYPE string,
             fillid            TYPE string,
             fontid            TYPE string,
             numfmtid          TYPE string,
             pivotbutton       TYPE string,
             quoteprefix       TYPE string,
             xfid              TYPE string,
           END OF lty_xf.

    TYPES: BEGIN OF lty_alignment,
             horizontal      TYPE string,
             indent          TYPE string,
             justifylastline TYPE string,
             readingorder    TYPE string,
             relativeindent  TYPE string,
             shrinktofit     TYPE string,
             textrotation    TYPE string,
             vertical        TYPE string,
             wraptext        TYPE string,
           END OF lty_alignment.

    TYPES: BEGIN OF lty_protection,
             hidden TYPE string,
             locked TYPE string,
           END OF lty_protection.

    DATA: lo_styles_xml                 TYPE REF TO if_ixml_document,
          lo_style                      TYPE REF TO zcl_excel_style,

          lt_num_formats                TYPE zcl_excel_style_number_format=>t_num_formats,
          lt_fills                      TYPE t_fills,
          lt_borders                    TYPE t_borders,
          lt_fonts                      TYPE t_fonts,

          ls_num_format                 TYPE zcl_excel_style_number_format=>t_num_format,
          ls_fill                       TYPE REF TO zcl_excel_style_fill,
          ls_cell_border                TYPE REF TO zcl_excel_style_borders,
          ls_font                       TYPE REF TO zcl_excel_style_font,

          lo_node_cellxfs               TYPE REF TO if_ixml_element,
          lo_node_cellxfs_xf            TYPE REF TO if_ixml_element,
          lo_node_cellxfs_xf_alignment  TYPE REF TO if_ixml_element,
          lo_node_cellxfs_xf_protection TYPE REF TO if_ixml_element,

          lo_nodes_xf                   TYPE REF TO if_ixml_node_collection,
          lo_iterator_cellxfs           TYPE REF TO if_ixml_node_iterator,

          ls_xf                         TYPE lty_xf,
          ls_alignment                  TYPE lty_alignment,
          ls_protection                 TYPE lty_protection,
          lv_index                      TYPE i.

*--------------------------------------------------------------------*
* To build a complete style that fully describes how a cell looks like
* we need the various parts
* §1 - Numberformat
* §2 - Fillstyle
* §3 - Borders
* §4 - Font
* §5 - Alignment
* §6 - Protection

*          Following is an example how this part of a file could be set up
*              ...
*              parts with various formatinformation - see §1,§2,§3,§4
*              ...
*          <cellXfs count="26">
*              <xf numFmtId="0" borderId="0" fillId="0" fontId="0" xfId="0"/>
*              <xf numFmtId="0" borderId="0" fillId="2" fontId="0" xfId="0" applyFill="1"/>
*              <xf numFmtId="0" borderId="1" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              <xf numFmtId="0" borderId="2" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              <xf numFmtId="0" borderId="3" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              <xf numFmtId="0" borderId="4" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              <xf numFmtId="0" borderId="0" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              ...
*          </cellXfs>
*--------------------------------------------------------------------*

    lo_styles_xml = me->get_ixml_from_zip_archive( ip_path ).

*--------------------------------------------------------------------*
* The styles are build up from
* §1 number formats
* §2 fill styles
* §3 border styles
* §4 fonts
* These need to be read before we can try to build up a complete
* style that describes the look of a cell
*--------------------------------------------------------------------*
    lt_num_formats   = load_style_num_formats( lo_styles_xml ).   " §1
    lt_fills         = load_style_fills( lo_styles_xml ).         " §2
    lt_borders       = load_style_borders( lo_styles_xml ).       " §3
    lt_fonts         = load_style_fonts( lo_styles_xml ).         " §4

*--------------------------------------------------------------------*
* Now everything is prepared to build a "full" style
*--------------------------------------------------------------------*
    lo_node_cellxfs  = lo_styles_xml->find_from_name_ns( name = 'cellXfs' uri = namespace-main ).
    IF lo_node_cellxfs IS BOUND.
      lo_nodes_xf         = lo_node_cellxfs->get_elements_by_tag_name_ns( name = 'xf' uri = namespace-main ).
      lo_iterator_cellxfs = lo_nodes_xf->create_iterator( ).
      lo_node_cellxfs_xf ?= lo_iterator_cellxfs->get_next( ).
      WHILE lo_node_cellxfs_xf IS BOUND.

        lo_style = ip_excel->add_new_style( ).
        fill_struct_from_attributes( EXPORTING
                                       ip_element   =  lo_node_cellxfs_xf
                                     CHANGING
                                       cp_structure = ls_xf ).
*--------------------------------------------------------------------*
* §2 fill style
*--------------------------------------------------------------------*
        IF ls_xf-applyfill = '1' AND ls_xf-fillid IS NOT INITIAL.
          lv_index = ls_xf-fillid + 1.
          READ TABLE lt_fills INTO ls_fill INDEX lv_index.
          IF sy-subrc = 0.
            lo_style->fill = ls_fill.
          ENDIF.
        ENDIF.

*--------------------------------------------------------------------*
* §1 number format
*--------------------------------------------------------------------*
        IF ls_xf-numfmtid IS NOT INITIAL.
          READ TABLE lt_num_formats INTO ls_num_format WITH TABLE KEY id = ls_xf-numfmtid.
          IF sy-subrc = 0.
            lo_style->number_format = ls_num_format-format.
          ENDIF.
        ENDIF.

*--------------------------------------------------------------------*
* §3 border style
*--------------------------------------------------------------------*
        IF ls_xf-applyborder = '1' AND ls_xf-borderid IS NOT INITIAL.
          lv_index = ls_xf-borderid + 1.
          READ TABLE lt_borders INTO ls_cell_border INDEX lv_index.
          IF sy-subrc = 0.
            lo_style->borders = ls_cell_border.
          ENDIF.
        ENDIF.

*--------------------------------------------------------------------*
* §4 font
*--------------------------------------------------------------------*
        IF ls_xf-applyfont = '1' AND ls_xf-fontid IS NOT INITIAL.
          lv_index = ls_xf-fontid + 1.
          READ TABLE lt_fonts INTO ls_font INDEX lv_index.
          IF sy-subrc = 0.
            lo_style->font = ls_font.
          ENDIF.
        ENDIF.

*--------------------------------------------------------------------*
* §5 - Alignment
*--------------------------------------------------------------------*
        lo_node_cellxfs_xf_alignment ?= lo_node_cellxfs_xf->find_from_name_ns( name = 'alignment' uri = namespace-main ).
        IF lo_node_cellxfs_xf_alignment IS BOUND.
          fill_struct_from_attributes( EXPORTING
                                         ip_element   =  lo_node_cellxfs_xf_alignment
                                       CHANGING
                                         cp_structure = ls_alignment ).
          IF ls_alignment-horizontal IS NOT INITIAL.
            lo_style->alignment->horizontal = ls_alignment-horizontal.
          ENDIF.

          IF ls_alignment-vertical IS NOT INITIAL.
            lo_style->alignment->vertical = ls_alignment-vertical.
          ENDIF.

          IF ls_alignment-textrotation IS NOT INITIAL.
            lo_style->alignment->textrotation = ls_alignment-textrotation.
          ENDIF.

          IF ls_alignment-wraptext = '1' OR ls_alignment-wraptext = 'true'.
            lo_style->alignment->wraptext = abap_true.
          ENDIF.

          IF ls_alignment-shrinktofit = '1' OR ls_alignment-shrinktofit = 'true'.
            lo_style->alignment->shrinktofit = abap_true.
          ENDIF.

          IF ls_alignment-indent IS NOT INITIAL.
            lo_style->alignment->indent = ls_alignment-indent.
          ENDIF.
        ENDIF.

*--------------------------------------------------------------------*
* §6 - Protection
*--------------------------------------------------------------------*
        lo_node_cellxfs_xf_protection ?= lo_node_cellxfs_xf->find_from_name_ns( name = 'protection' uri = namespace-main ).
        IF lo_node_cellxfs_xf_protection IS BOUND.
          fill_struct_from_attributes( EXPORTING
                                         ip_element   = lo_node_cellxfs_xf_protection
                                       CHANGING
                                         cp_structure = ls_protection ).
          IF ls_protection-locked = '1' OR ls_protection-locked = 'true'.
            lo_style->protection->locked = zcl_excel_style_protection=>c_protection_locked.
          ELSE.
            lo_style->protection->locked = zcl_excel_style_protection=>c_protection_unlocked.
          ENDIF.

          IF ls_protection-hidden = '1' OR ls_protection-hidden = 'true'.
            lo_style->protection->hidden = zcl_excel_style_protection=>c_protection_hidden.
          ELSE.
            lo_style->protection->hidden = zcl_excel_style_protection=>c_protection_unhidden.
          ENDIF.

        ENDIF.

        INSERT lo_style INTO TABLE me->styles.

        lo_node_cellxfs_xf ?= lo_iterator_cellxfs->get_next( ).

      ENDWHILE.
    ENDIF.

  ENDMETHOD.