  METHOD zif_excel_style_changer~set_complete_alignment.

    DATA: alignmentx LIKE ip_xalignment.

    IF ip_xalignment IS SUPPLIED.
      alignmentx = ip_xalignment.
    ELSE.
      CLEAR alignmentx WITH 'X'.
      IF ip_alignment-horizontal IS INITIAL.
        CLEAR alignmentx-horizontal.
      ENDIF.
      IF ip_alignment-vertical IS INITIAL.
        CLEAR alignmentx-vertical.
      ENDIF.
    ENDIF.

    complete_style-alignment  = ip_alignment .
    complete_stylex-alignment = alignmentx   .
    multiple_change_requested-alignment = abap_true.

    result = me.

  ENDMETHOD.