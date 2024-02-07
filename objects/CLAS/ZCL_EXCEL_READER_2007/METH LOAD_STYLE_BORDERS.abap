  METHOD load_style_borders.

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          aligning code
*          renaming variables to indicate what they are used for
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
    DATA: lo_node_border      TYPE REF TO if_ixml_element,
          lo_node_bordertype  TYPE REF TO if_ixml_element,
          lo_node_bordercolor TYPE REF TO if_ixml_element,
          lo_cell_border      TYPE REF TO zcl_excel_style_borders,
          lo_border           TYPE REF TO zcl_excel_style_border,
          ls_color            TYPE t_color.

*--------------------------------------------------------------------*
* We need a table of used borderformats to build up our styles
* §1    A cell has 4 outer borders and 2 diagonal "borders"
*       These borders can be formatted separately but the diagonal borders
*       are always being formatted the same
*       We'll parse through the <border>-tag for each of the bordertypes
* §2    and read the corresponding formatting information

*          Following is an example how this part of a file could be set up
*          <border diagonalDown="1">
*              <left style="mediumDashDotDot">
*                  <color rgb="FFFF0000"/>
*              </left>
*              <right/>
*              <top style="thick">
*                  <color rgb="FFFF0000"/>
*              </top>
*              <bottom style="thick">
*                  <color rgb="FFFF0000"/>
*              </bottom>
*              <diagonal style="thick">
*                  <color rgb="FFFF0000"/>
*              </diagonal>
*          </border>
*--------------------------------------------------------------------*
    lo_node_border ?= ip_xml->find_from_name_ns( name = 'border' uri = namespace-main ).
    WHILE lo_node_border IS BOUND.

      CREATE OBJECT lo_cell_border.

*--------------------------------------------------------------------*
* Diagonal borderlines are formatted the equally.  Determine what kind of diagonal borders are present if any
*--------------------------------------------------------------------*
* DiagonalNone = 0
* DiagonalUp   = 1
* DiagonalDown = 2
* DiagonalBoth = 3
*--------------------------------------------------------------------*
      IF lo_node_border->get_attribute( 'diagonalDown' ) IS NOT INITIAL.
        ADD zcl_excel_style_borders=>c_diagonal_down TO lo_cell_border->diagonal_mode.
      ENDIF.

      IF lo_node_border->get_attribute( 'diagonalUp' ) IS NOT INITIAL.
        ADD zcl_excel_style_borders=>c_diagonal_up TO lo_cell_border->diagonal_mode.
      ENDIF.

      lo_node_bordertype ?= lo_node_border->get_first_child( ).
      WHILE lo_node_bordertype IS BOUND.
*--------------------------------------------------------------------*
* §1 Determine what kind of border we are talking about
*--------------------------------------------------------------------*
* Up, down, left, right, diagonal
*--------------------------------------------------------------------*
        CREATE OBJECT lo_border.

        CASE lo_node_bordertype->get_name( ).

          WHEN 'left'.
            lo_cell_border->left = lo_border.

          WHEN 'right'.
            lo_cell_border->right = lo_border.

          WHEN 'top'.
            lo_cell_border->top = lo_border.

          WHEN 'bottom'.
            lo_cell_border->down = lo_border.

          WHEN 'diagonal'.
            lo_cell_border->diagonal = lo_border.

        ENDCASE.

*--------------------------------------------------------------------*
* §2 Read the border-formatting
*--------------------------------------------------------------------*
        lo_border->border_style = lo_node_bordertype->get_attribute( 'style' ).
        lo_node_bordercolor ?= lo_node_bordertype->find_from_name_ns( name = 'color' uri = namespace-main ).
        IF lo_node_bordercolor IS BOUND.
          fill_struct_from_attributes( EXPORTING
                                         ip_element   =  lo_node_bordercolor
                                       CHANGING
                                         cp_structure = ls_color ).

          lo_border->border_color-rgb = ls_color-rgb.
          IF ls_color-indexed IS NOT INITIAL.
            lo_border->border_color-indexed = ls_color-indexed.
          ENDIF.

          IF ls_color-theme IS NOT INITIAL.
            lo_border->border_color-theme = ls_color-theme.
          ENDIF.
          lo_border->border_color-tint = ls_color-tint.
        ENDIF.

        lo_node_bordertype ?= lo_node_bordertype->get_next( ).

      ENDWHILE.

      INSERT lo_cell_border INTO TABLE ep_borders.

      lo_node_border ?= lo_node_border->get_next( ).

    ENDWHILE.


  ENDMETHOD.