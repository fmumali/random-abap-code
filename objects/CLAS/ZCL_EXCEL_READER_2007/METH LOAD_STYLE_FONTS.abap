  METHOD load_style_fonts.

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          aligning code
*          removing unused variables
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
    DATA: lo_node_font TYPE REF TO if_ixml_element,
          lo_font      TYPE REF TO zcl_excel_style_font.

*--------------------------------------------------------------------*
* We need a table of used fonts to build up our styles

*          Following is an example how this part of a file could be set up
*          <font>
*              <sz val="11"/>
*              <color theme="1"/>
*              <name val="Calibri"/>
*              <family val="2"/>
*              <scheme val="minor"/>
*          </font>
*--------------------------------------------------------------------*
    lo_node_font ?= ip_xml->find_from_name_ns( name = 'font' uri = namespace-main ).
    WHILE lo_node_font IS BOUND.

      lo_font = load_style_font( lo_node_font ).
      INSERT lo_font INTO TABLE ep_fonts.

      lo_node_font ?= lo_node_font->get_next( ).

    ENDWHILE.


  ENDMETHOD.