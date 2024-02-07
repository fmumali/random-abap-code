  METHOD unescape_string_value.

    DATA:
      "Marks the Position before the searched Pattern occurs in the String
      "For example in String A_X_TEST_X, the Table is filled with 1 and 8
      lt_character_positions       TYPE TABLE OF i,
      lv_character_position        TYPE i,
      lv_character_position_plus_2 TYPE i,
      lv_character_position_plus_6 TYPE i,
      lv_unescaped_value           TYPE string.

    " The text "_x...._", with "_x" not "_X". Each "." represents one character, being 0-9 a-f or A-F (case insensitive),
    " is interpreted like Unicode character U+.... (e.g. "_x0041_" is rendered like "A") is for characters.
    " To not interpret it, Excel replaces the first "_" with "_x005f_".
    result = i_value.

    IF provided_string_is_escaped( i_value ) = abap_true.
      CLEAR lt_character_positions.
      APPEND sy-fdpos TO lt_character_positions.
      lv_character_position = sy-fdpos + 1.
      WHILE result+lv_character_position CS '_x'.
        ADD sy-fdpos TO lv_character_position.
        APPEND lv_character_position TO lt_character_positions.
        ADD 1 TO lv_character_position.
      ENDWHILE.
      SORT lt_character_positions BY table_line DESCENDING.
      LOOP AT lt_character_positions INTO lv_character_position.
        lv_character_position_plus_2 = lv_character_position + 2.
        lv_character_position_plus_6 = lv_character_position + 6.
        IF substring( val = result off = lv_character_position_plus_2 len = 4 ) CO '0123456789ABCDEFabcdef'.
          IF substring( val = result off = lv_character_position_plus_6 len = 1 ) = '_'.
            lv_unescaped_value = cl_abap_conv_in_ce=>uccp( to_upper( substring( val = result off = lv_character_position_plus_2 len = 4 ) ) ).
            REPLACE SECTION OFFSET lv_character_position LENGTH 7 OF result WITH lv_unescaped_value.
          ENDIF.
        ENDIF.
      ENDLOOP.
    ENDIF.

  ENDMETHOD.