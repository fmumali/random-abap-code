  METHOD escape_string_value.

    DATA: lt_character_positions TYPE TABLE OF i,
          lv_character_position  TYPE i,
          lv_escaped_value       TYPE string.

    result = iv_value.
    IF result CA control_characters.

      CLEAR lt_character_positions.
      APPEND sy-fdpos TO lt_character_positions.
      lv_character_position = sy-fdpos + 1.
      WHILE result+lv_character_position CA control_characters.
        ADD sy-fdpos TO lv_character_position.
        APPEND lv_character_position TO lt_character_positions.
        ADD 1 TO lv_character_position.
      ENDWHILE.
      SORT lt_character_positions BY table_line DESCENDING.

      LOOP AT lt_character_positions INTO lv_character_position.
        lv_escaped_value = |_x{ cl_abap_conv_out_ce=>uccp( substring( val = result off = lv_character_position len = 1 ) ) }_|.
        REPLACE SECTION OFFSET lv_character_position LENGTH 1 OF result WITH lv_escaped_value.
      ENDLOOP.
    ENDIF.

  ENDMETHOD.