  METHOD clean_fieldcatalog.
    DATA: l_position TYPE tabfdpos.

    FIELD-SYMBOLS: <fs_sfcat>   TYPE zexcel_s_converter_fcat.

    SORT wt_fieldcatalog BY position col_id.

    CLEAR l_position.
    LOOP AT wt_fieldcatalog ASSIGNING <fs_sfcat>.
      ADD 1 TO l_position.
      <fs_sfcat>-position = l_position.
* Default stype with alignment and format
      <fs_sfcat>-style_hdr      = get_style( i_type      = c_type_hdr
                                             i_alignment = <fs_sfcat>-alignment ).
      IF ws_layout-is_stripped = abap_true.
        <fs_sfcat>-style_stripped = get_style( i_type      = c_type_str
                                               i_alignment = <fs_sfcat>-alignment
                                               i_inttype   = <fs_sfcat>-inttype
                                               i_decimals  = <fs_sfcat>-decimals   ).
      ENDIF.
      <fs_sfcat>-style_normal   = get_style( i_type      = c_type_nor
                                             i_alignment = <fs_sfcat>-alignment
                                             i_inttype   = <fs_sfcat>-inttype
                                             i_decimals  = <fs_sfcat>-decimals   ).
      <fs_sfcat>-style_subtotal = get_style( i_type      = c_type_sub
                                             i_alignment = <fs_sfcat>-alignment
                                             i_inttype   = <fs_sfcat>-inttype
                                             i_decimals  = <fs_sfcat>-decimals   ).
      <fs_sfcat>-style_total    = get_style( i_type      = c_type_tot
                                             i_alignment = <fs_sfcat>-alignment
                                             i_inttype   = <fs_sfcat>-inttype
                                             i_decimals  = <fs_sfcat>-decimals   ).
    ENDLOOP.

  ENDMETHOD.