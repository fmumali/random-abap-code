  METHOD get_color.
    DATA: ls_con_col TYPE zexcel_s_converter_col,
          ls_color   TYPE ts_col_converter,
          l_line     TYPE i,
          l_color(4) TYPE c.
    FIELD-SYMBOLS: <fs_tab>  TYPE STANDARD TABLE,
                   <fs_stab> TYPE any,
                   <fs>      TYPE any,
                   <ft_slis> TYPE STANDARD TABLE,
                   <fs_slis> TYPE any.

* Loop trough the table to set the color properties of each line. The color properties field is
* Char 4 and the characters is set as follows:
* Char 1 = C = This is a color property
* Char 2 = 6 = Color code (1 - 7)
* Char 3 = Intensified on/of = 1 = on
* Char 4 = Inverse display = 0 = of

    ASSIGN io_table->* TO <fs_tab>.

    IF ws_layo-info_fname IS NOT INITIAL OR
       ws_layo-ctab_fname IS NOT INITIAL.
      LOOP AT <fs_tab> ASSIGNING <fs_stab>.
        l_line = sy-tabix.
        IF ws_layo-info_fname IS NOT INITIAL.
          ASSIGN COMPONENT ws_layo-info_fname OF STRUCTURE <fs_stab> TO <fs>.
          IF sy-subrc = 0 AND <fs> IS NOT INITIAL.
            l_color = <fs>.
            IF l_color(1) = 'C'.
              READ TABLE wt_colors INTO ls_color WITH TABLE KEY col = l_color+1(1)
                                                                int = l_color+2(1)
                                                                inv = l_color+3(1).
              IF sy-subrc = 0.
                ls_con_col-rownumber  = l_line.
                ls_con_col-columnname = space.
                ls_con_col-fontcolor  = ls_color-fontcolor.
                ls_con_col-fillcolor  = ls_color-fillcolor.
                INSERT ls_con_col INTO TABLE et_colors.
              ENDIF.
            ENDIF.
          ENDIF.
        ENDIF.
        IF ws_layo-ctab_fname IS NOT INITIAL.

          ASSIGN COMPONENT ws_layo-ctab_fname OF STRUCTURE <fs_stab> TO <ft_slis>.
          IF sy-subrc = 0.
            LOOP AT <ft_slis> ASSIGNING <fs_slis>.
              ASSIGN COMPONENT 'COLOR' OF STRUCTURE <fs_slis> TO <fs>.
              IF sy-subrc = 0.
                IF <fs> IS NOT INITIAL.
                  FIELD-SYMBOLS: <col>      TYPE any,
                                 <int>      TYPE any,
                                 <inv>      TYPE any,
                                 <fname>    TYPE any,
                                 <nokeycol> TYPE any.
                  ASSIGN COMPONENT 'COL' OF STRUCTURE <fs> TO <col>.
                  ASSIGN COMPONENT 'INT' OF STRUCTURE <fs> TO <int>.
                  ASSIGN COMPONENT 'INV' OF STRUCTURE <fs> TO <inv>.
                  READ TABLE wt_colors INTO ls_color WITH TABLE KEY col = <col>
                                                                    int = <int>
                                                                    inv = <inv>.
                  IF sy-subrc = 0.
                    ls_con_col-rownumber  = l_line.
                    ASSIGN COMPONENT 'FNAME' OF STRUCTURE <fs_slis> TO <fname>.
                    IF sy-subrc NE 0.
                      ASSIGN COMPONENT 'FIELDNAME' OF STRUCTURE <fs_slis> TO <fname>.
                      IF sy-subrc EQ 0.
                        ls_con_col-columnname = <fname>.
                      ENDIF.
                    ELSE.
                      ls_con_col-columnname = <fname>.
                    ENDIF.

                    ls_con_col-fontcolor  = ls_color-fontcolor.
                    ls_con_col-fillcolor  = ls_color-fillcolor.
                    ASSIGN COMPONENT 'NOKEYCOL' OF STRUCTURE <fs_slis> TO <nokeycol>.
                    IF sy-subrc EQ 0.
                      ls_con_col-nokeycol   = <nokeycol>.
                    ENDIF.
                    INSERT ls_con_col INTO TABLE et_colors.
                  ENDIF.
                ENDIF.
              ENDIF.
            ENDLOOP.
          ENDIF.
        ENDIF.
      ENDLOOP.
    ENDIF.
  ENDMETHOD.