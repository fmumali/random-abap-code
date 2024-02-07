  METHOD load_shared_strings.

    DATA: lo_reader TYPE REF TO if_sxml_reader.
    DATA: lt_shared_strings TYPE TABLE OF string,
          ls_shared_string  TYPE t_shared_string.
    FIELD-SYMBOLS: <lv_shared_string> TYPE string.

    lo_reader = get_sxml_reader( ip_path ).

    lt_shared_strings = read_shared_strings( lo_reader ).
    LOOP AT lt_shared_strings ASSIGNING <lv_shared_string>.
      ls_shared_string-value = unescape_string_value( <lv_shared_string> ).
      APPEND ls_shared_string TO shared_strings.
    ENDLOOP.

  ENDMETHOD.