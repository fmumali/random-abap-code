  METHOD load_style_fills.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Support gradientFill
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          aligning code
*          commenting on problems/future enhancements/todos we already know of or should decide upon
*          adding comments to explain what we are trying to achieve
*          renaming variables to indicate what they are used for
*--------------------------------------------------------------------*
    DATA: lv_value           TYPE string,
          lo_node_fill       TYPE REF TO if_ixml_element,
          lo_node_fill_child TYPE REF TO if_ixml_element,
          lo_node_bgcolor    TYPE REF TO if_ixml_element,
          lo_node_fgcolor    TYPE REF TO if_ixml_element,
          lo_node_stop       TYPE REF TO if_ixml_element,
          lo_fill            TYPE REF TO zcl_excel_style_fill,
          ls_color           TYPE t_color.

*--------------------------------------------------------------------*
* We need a table of used fillformats to build up our styles

*          Following is an example how this part of a file could be set up
*          <fill>
*              <patternFill patternType="gray125"/>
*          </fill>
*          <fill>
*              <patternFill patternType="solid">
*                  <fgColor rgb="FFFFFF00"/>
*                  <bgColor indexed="64"/>
*              </patternFill>
*          </fill>
*--------------------------------------------------------------------*

    lo_node_fill ?= ip_xml->find_from_name_ns( name = 'fill' uri = namespace-main ).
    WHILE lo_node_fill IS BOUND.

      CREATE OBJECT lo_fill.
      lo_node_fill_child ?= lo_node_fill->get_first_child( ).
      lv_value            = lo_node_fill_child->get_name( ).
      CASE lv_value.

*--------------------------------------------------------------------*
* Patternfill
*--------------------------------------------------------------------*
        WHEN 'patternFill'.
          lo_fill->filltype = lo_node_fill_child->get_attribute( 'patternType' ).
*--------------------------------------------------------------------*
* Patternfill - background color
*--------------------------------------------------------------------*
          lo_node_bgcolor = lo_node_fill_child->find_from_name_ns( name = 'bgColor' uri = namespace-main ).
          IF lo_node_bgcolor IS BOUND.
            fill_struct_from_attributes( EXPORTING
                                           ip_element   = lo_node_bgcolor
                                         CHANGING
                                           cp_structure = ls_color ).

            lo_fill->bgcolor-rgb = ls_color-rgb.
            IF ls_color-indexed IS NOT INITIAL.
              lo_fill->bgcolor-indexed = ls_color-indexed.
            ENDIF.

            IF ls_color-theme IS NOT INITIAL.
              lo_fill->bgcolor-theme = ls_color-theme.
            ENDIF.
            lo_fill->bgcolor-tint = ls_color-tint.
          ENDIF.

*--------------------------------------------------------------------*
* Patternfill - foreground color
*--------------------------------------------------------------------*
          lo_node_fgcolor = lo_node_fill->find_from_name_ns( name = 'fgColor' uri = namespace-main ).
          IF lo_node_fgcolor IS BOUND.
            fill_struct_from_attributes( EXPORTING
                                           ip_element   = lo_node_fgcolor
                                         CHANGING
                                           cp_structure = ls_color ).

            lo_fill->fgcolor-rgb = ls_color-rgb.
            IF ls_color-indexed IS NOT INITIAL.
              lo_fill->fgcolor-indexed = ls_color-indexed.
            ENDIF.

            IF ls_color-theme IS NOT INITIAL.
              lo_fill->fgcolor-theme = ls_color-theme.
            ENDIF.
            lo_fill->fgcolor-tint = ls_color-tint.
          ENDIF.


*--------------------------------------------------------------------*
* gradientFill
*--------------------------------------------------------------------*
        WHEN 'gradientFill'.
          lo_fill->gradtype-type   = lo_node_fill_child->get_attribute( 'type' ).
          lo_fill->gradtype-top    = lo_node_fill_child->get_attribute( 'top' ).
          lo_fill->gradtype-left   = lo_node_fill_child->get_attribute( 'left' ).
          lo_fill->gradtype-right  = lo_node_fill_child->get_attribute( 'right' ).
          lo_fill->gradtype-bottom = lo_node_fill_child->get_attribute( 'bottom' ).
          lo_fill->gradtype-degree = lo_node_fill_child->get_attribute( 'degree' ).
          FREE lo_node_stop.
          lo_node_stop ?= lo_node_fill_child->find_from_name_ns( name = 'stop' uri = namespace-main ).
          WHILE lo_node_stop IS BOUND.
            IF lo_fill->gradtype-position1 IS INITIAL.
              lo_fill->gradtype-position1 = lo_node_stop->get_attribute( 'position' ).
              lo_node_bgcolor = lo_node_stop->find_from_name_ns( name = 'color' uri = namespace-main ).
              IF lo_node_bgcolor IS BOUND.
                fill_struct_from_attributes( EXPORTING
                                                ip_element   = lo_node_bgcolor
                                              CHANGING
                                                cp_structure = ls_color ).

                lo_fill->bgcolor-rgb = ls_color-rgb.
                IF ls_color-indexed IS NOT INITIAL.
                  lo_fill->bgcolor-indexed = ls_color-indexed.
                ENDIF.

                IF ls_color-theme IS NOT INITIAL.
                  lo_fill->bgcolor-theme = ls_color-theme.
                ENDIF.
                lo_fill->bgcolor-tint = ls_color-tint.
              ENDIF.
            ELSEIF lo_fill->gradtype-position2 IS INITIAL.
              lo_fill->gradtype-position2 = lo_node_stop->get_attribute( 'position' ).
              lo_node_fgcolor = lo_node_stop->find_from_name_ns( name = 'color' uri = namespace-main ).
              IF lo_node_fgcolor IS BOUND.
                fill_struct_from_attributes( EXPORTING
                                               ip_element   = lo_node_fgcolor
                                             CHANGING
                                               cp_structure = ls_color ).

                lo_fill->fgcolor-rgb = ls_color-rgb.
                IF ls_color-indexed IS NOT INITIAL.
                  lo_fill->fgcolor-indexed = ls_color-indexed.
                ENDIF.

                IF ls_color-theme IS NOT INITIAL.
                  lo_fill->fgcolor-theme = ls_color-theme.
                ENDIF.
                lo_fill->fgcolor-tint = ls_color-tint.
              ENDIF.
            ELSEIF lo_fill->gradtype-position3 IS INITIAL.
              lo_fill->gradtype-position3 = lo_node_stop->get_attribute( 'position' ).
              "BGColor is filled already with position 1 no need to check again
            ENDIF.

            lo_node_stop ?= lo_node_stop->get_next( ).
          ENDWHILE.

        WHEN OTHERS.

      ENDCASE.


      INSERT lo_fill INTO TABLE ep_fills.

      lo_node_fill ?= lo_node_fill->get_next( ).

    ENDWHILE.


  ENDMETHOD.