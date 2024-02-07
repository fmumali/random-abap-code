  METHOD get_dxf_style_guid.
    DATA: lo_ixml_dxf_children          TYPE REF TO if_ixml_node_list,
          lo_ixml_iterator_dxf_children TYPE REF TO if_ixml_node_iterator,
          lo_ixml_dxf_child             TYPE REF TO if_ixml_element,

          lv_dxf_child_type             TYPE string,

          lo_ixml_element               TYPE REF TO if_ixml_element,
          lo_ixml_element2              TYPE REF TO if_ixml_element,
          lv_val                        TYPE string.

    DATA: ls_cstyle  TYPE zexcel_s_cstyle_complete,
          ls_cstylex TYPE zexcel_s_cstylex_complete.



    lo_ixml_dxf_children = io_ixml_dxf->get_children( ).
    lo_ixml_iterator_dxf_children = lo_ixml_dxf_children->create_iterator( ).
    lo_ixml_dxf_child ?= lo_ixml_iterator_dxf_children->get_next( ).
    WHILE lo_ixml_dxf_child IS BOUND.

      lv_dxf_child_type = lo_ixml_dxf_child->get_name( ).
      CASE lv_dxf_child_type.

        WHEN 'font'.
*--------------------------------------------------------------------*
* italic
*--------------------------------------------------------------------*
          lo_ixml_element = lo_ixml_dxf_child->find_from_name_ns( name = 'i' uri = namespace-main ).
          IF lo_ixml_element IS BOUND.
            CLEAR lv_val.
            lv_val  = lo_ixml_element->get_attribute_ns( 'val' ).
            IF lv_val <> '0'.
              ls_cstyle-font-italic  = 'X'.
              ls_cstylex-font-italic = 'X'.
            ENDIF.

          ENDIF.
*--------------------------------------------------------------------*
* bold
*--------------------------------------------------------------------*
          lo_ixml_element = lo_ixml_dxf_child->find_from_name_ns( name = 'b' uri = namespace-main ).
          IF lo_ixml_element IS BOUND.
            CLEAR lv_val.
            lv_val  = lo_ixml_element->get_attribute_ns( 'val' ).
            IF lv_val <> '0'.
              ls_cstyle-font-bold  = 'X'.
              ls_cstylex-font-bold = 'X'.
            ENDIF.

          ENDIF.
*--------------------------------------------------------------------*
* strikethrough
*--------------------------------------------------------------------*
          lo_ixml_element = lo_ixml_dxf_child->find_from_name_ns( name = 'strike' uri = namespace-main ).
          IF lo_ixml_element IS BOUND.
            CLEAR lv_val.
            lv_val  = lo_ixml_element->get_attribute_ns( 'val' ).
            IF lv_val <> '0'.
              ls_cstyle-font-strikethrough  = 'X'.
              ls_cstylex-font-strikethrough = 'X'.
            ENDIF.

          ENDIF.
*--------------------------------------------------------------------*
* color
*--------------------------------------------------------------------*
          lo_ixml_element = lo_ixml_dxf_child->find_from_name_ns( name = 'color' uri = namespace-main ).
          IF lo_ixml_element IS BOUND.
            CLEAR lv_val.
            lv_val  = lo_ixml_element->get_attribute_ns( 'rgb' ).
            ls_cstyle-font-color-rgb  = lv_val.
            ls_cstylex-font-color-rgb = 'X'.
          ENDIF.

        WHEN 'fill'.
          lo_ixml_element = lo_ixml_dxf_child->find_from_name_ns( name = 'patternFill' uri = namespace-main ).
          IF lo_ixml_element IS BOUND.
            lo_ixml_element2 = lo_ixml_dxf_child->find_from_name_ns( name = 'bgColor' uri = namespace-main ).
            IF lo_ixml_element2 IS BOUND.
              CLEAR lv_val.
              lv_val  = lo_ixml_element2->get_attribute_ns( 'rgb' ).
              IF lv_val IS NOT INITIAL.
                ls_cstyle-fill-filltype       = zcl_excel_style_fill=>c_fill_solid.
                ls_cstyle-fill-bgcolor-rgb    = lv_val.
                ls_cstylex-fill-filltype      = 'X'.
                ls_cstylex-fill-bgcolor-rgb   = 'X'.
              ENDIF.
              CLEAR lv_val.
              lv_val  = lo_ixml_element2->get_attribute_ns( 'theme' ).
              IF lv_val IS NOT INITIAL.
                ls_cstyle-fill-filltype         = zcl_excel_style_fill=>c_fill_solid.
                ls_cstyle-fill-bgcolor-theme    = lv_val.
                ls_cstylex-fill-filltype        = 'X'.
                ls_cstylex-fill-bgcolor-theme   = 'X'.
              ENDIF.
              CLEAR lv_val.
            ENDIF.
          ENDIF.

      ENDCASE.

      lo_ixml_dxf_child ?= lo_ixml_iterator_dxf_children->get_next( ).

    ENDWHILE.

    rv_style_guid = io_excel->get_static_cellstyle_guid( ip_cstyle_complete  = ls_cstyle
                                                         ip_cstylex_complete = ls_cstylex  ).


  ENDMETHOD.