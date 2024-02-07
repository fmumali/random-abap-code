  METHOD load.
    DATA: lo_scheme_children TYPE REF TO if_ixml_node_list.
    DATA: lo_scheme_iterator TYPE REF TO if_ixml_node_iterator.
    DATA: lo_scheme_element TYPE REF TO if_ixml_element.
    DATA: lo_major_children TYPE REF TO if_ixml_node_list.
    DATA: lo_major_iterator TYPE REF TO if_ixml_node_iterator.
    DATA: lo_major_element TYPE REF TO if_ixml_element.
    DATA: lo_minor_children TYPE REF TO if_ixml_node_list.
    DATA: lo_minor_iterator TYPE REF TO if_ixml_node_iterator.
    DATA: lo_minor_element TYPE REF TO if_ixml_element.
    DATA: ls_font TYPE t_font.
    CHECK io_font_scheme IS NOT INITIAL.
    CLEAR font_scheme.
    font_scheme-name =  io_font_scheme->get_attribute( name = c_name ).
    lo_scheme_children = io_font_scheme->get_children( ).
    lo_scheme_iterator = lo_scheme_children->create_iterator( ).
    lo_scheme_element ?= lo_scheme_iterator->get_next( ).
    WHILE lo_scheme_element IS BOUND.
      CASE lo_scheme_element->get_name( ).
        WHEN c_major.
          lo_major_children = lo_scheme_element->get_children( ).
          lo_major_iterator = lo_major_children->create_iterator( ).
          lo_major_element ?= lo_major_iterator->get_next( ).
          WHILE lo_major_element IS BOUND.
            CASE lo_major_element->get_name( ).
              WHEN c_latin.
                font_scheme-major-latin-typeface = lo_major_element->get_attribute(  name = c_typeface ).
                font_scheme-major-latin-panose = lo_major_element->get_attribute(  name = c_panose ).
                font_scheme-major-latin-pitchfamily = lo_major_element->get_attribute(  name = c_pitchfamily ).
                font_scheme-major-latin-charset = lo_major_element->get_attribute(  name = c_charset ).
              WHEN c_ea.
                font_scheme-major-ea-typeface = lo_major_element->get_attribute(  name = c_typeface ).
                font_scheme-major-ea-panose = lo_major_element->get_attribute(  name = c_panose ).
                font_scheme-major-ea-pitchfamily = lo_major_element->get_attribute(  name = c_pitchfamily ).
                font_scheme-major-ea-charset = lo_major_element->get_attribute(  name = c_charset ).
              WHEN c_cs.
                font_scheme-major-cs-typeface = lo_major_element->get_attribute(  name = c_typeface ).
                font_scheme-major-cs-panose = lo_major_element->get_attribute(  name = c_panose ).
                font_scheme-major-cs-pitchfamily = lo_major_element->get_attribute(  name = c_pitchfamily ).
                font_scheme-major-cs-charset = lo_major_element->get_attribute(  name = c_charset ).
              WHEN c_font.
                CLEAR ls_font.
                ls_font-script = lo_major_element->get_attribute(  name = c_script ).
                ls_font-typeface = lo_major_element->get_attribute(  name = c_typeface ).
                TRY.
                    INSERT ls_font INTO TABLE font_scheme-major-fonts.
                  CATCH cx_root. "not the best but just to avoid duplicate lines dump

                ENDTRY.
            ENDCASE.
            lo_major_element ?= lo_major_iterator->get_next( ).
          ENDWHILE.
        WHEN c_minor.
          lo_minor_children = lo_scheme_element->get_children( ).
          lo_minor_iterator = lo_minor_children->create_iterator( ).
          lo_minor_element ?= lo_minor_iterator->get_next( ).
          WHILE lo_minor_element IS BOUND.
            CASE lo_minor_element->get_name( ).
              WHEN c_latin.
                font_scheme-minor-latin-typeface = lo_minor_element->get_attribute(  name = c_typeface ).
                font_scheme-minor-latin-panose = lo_minor_element->get_attribute(  name = c_panose ).
                font_scheme-minor-latin-pitchfamily = lo_minor_element->get_attribute(  name = c_pitchfamily ).
                font_scheme-minor-latin-charset = lo_minor_element->get_attribute(  name = c_charset ).
              WHEN c_ea.
                font_scheme-minor-ea-typeface = lo_minor_element->get_attribute(  name = c_typeface ).
                font_scheme-minor-ea-panose = lo_minor_element->get_attribute(  name = c_panose ).
                font_scheme-minor-ea-pitchfamily = lo_minor_element->get_attribute(  name = c_pitchfamily ).
                font_scheme-minor-ea-charset = lo_minor_element->get_attribute(  name = c_charset ).
              WHEN c_cs.
                font_scheme-minor-cs-typeface = lo_minor_element->get_attribute(  name = c_typeface ).
                font_scheme-minor-cs-panose = lo_minor_element->get_attribute(  name = c_panose ).
                font_scheme-minor-cs-pitchfamily = lo_minor_element->get_attribute(  name = c_pitchfamily ).
                font_scheme-minor-cs-charset = lo_minor_element->get_attribute(  name = c_charset ).
              WHEN c_font.
                CLEAR ls_font.
                ls_font-script = lo_minor_element->get_attribute(  name = c_script ).
                ls_font-typeface = lo_minor_element->get_attribute(  name = c_typeface ).
                TRY.
                    INSERT ls_font INTO TABLE font_scheme-minor-fonts.
                  CATCH cx_root. "not the best but just to avoid duplicate lines dump

                ENDTRY.
            ENDCASE.
            lo_minor_element ?= lo_minor_iterator->get_next( ).
          ENDWHILE.
      ENDCASE.
      lo_scheme_element ?= lo_scheme_iterator->get_next( ).
    ENDWHILE.
  ENDMETHOD.                    "load