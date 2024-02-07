  METHOD get_header_footer_string.
* ----------------------------------------------------------------------
    DATA:   lc_marker_left(2)   TYPE c VALUE '&L'
          , lc_marker_right(2)  TYPE c VALUE '&R'
          , lc_marker_center(2) TYPE c VALUE '&C'
          , lv_value            TYPE string
          .
* ----------------------------------------------------------------------
    IF ep_odd_header IS SUPPLIED.

      IF me->odd_header-left_value IS NOT INITIAL.
        lv_value = me->process_header_footer( ip_header = me->odd_header ip_side = 'LEFT' ).
        CONCATENATE lc_marker_left lv_value INTO ep_odd_header.
      ENDIF.

      IF me->odd_header-center_value IS NOT INITIAL.
        lv_value = me->process_header_footer( ip_header = me->odd_header ip_side = 'CENTER' ).
        CONCATENATE ep_odd_header lc_marker_center lv_value INTO ep_odd_header.
      ENDIF.

      IF me->odd_header-right_value IS NOT INITIAL.
        lv_value = me->process_header_footer( ip_header = me->odd_header ip_side = 'RIGHT' ).
        CONCATENATE ep_odd_header lc_marker_right lv_value INTO ep_odd_header.
      ENDIF.

      IF me->odd_header-left_image IS NOT INITIAL.
        lv_value = '&G'.
        CONCATENATE ep_odd_header lc_marker_left lv_value INTO ep_odd_header.
      ENDIF.
      IF me->odd_header-center_image IS NOT INITIAL.
        lv_value = '&G'.
        CONCATENATE ep_odd_header lc_marker_center lv_value INTO ep_odd_header.
      ENDIF.
      IF me->odd_header-right_image IS NOT INITIAL.
        lv_value = '&G'.
        CONCATENATE ep_odd_header lc_marker_right lv_value INTO ep_odd_header.
      ENDIF.

    ENDIF.
* ----------------------------------------------------------------------
    IF ep_odd_footer IS SUPPLIED.

      IF me->odd_footer-left_value IS NOT INITIAL.
        lv_value = me->process_header_footer( ip_header = me->odd_footer ip_side = 'LEFT' ).
        CONCATENATE lc_marker_left lv_value INTO ep_odd_footer.
      ENDIF.

      IF me->odd_footer-center_value IS NOT INITIAL.
        lv_value = me->process_header_footer( ip_header = me->odd_footer ip_side = 'CENTER' ).
        CONCATENATE ep_odd_footer lc_marker_center lv_value INTO ep_odd_footer.
      ENDIF.

      IF me->odd_footer-right_value IS NOT INITIAL.
        lv_value = me->process_header_footer( ip_header = me->odd_footer ip_side = 'RIGHT' ).
        CONCATENATE ep_odd_footer lc_marker_right lv_value INTO ep_odd_footer.
      ENDIF.

      IF me->odd_footer-left_image IS NOT INITIAL.
        lv_value = '&G'.
        CONCATENATE ep_odd_header lc_marker_left lv_value INTO ep_odd_footer.
      ENDIF.
      IF me->odd_footer-center_image IS NOT INITIAL.
        lv_value = '&G'.
        CONCATENATE ep_odd_header lc_marker_center lv_value INTO ep_odd_footer.
      ENDIF.
      IF me->odd_footer-right_image IS NOT INITIAL.
        lv_value = '&G'.
        CONCATENATE ep_odd_header lc_marker_right lv_value INTO ep_odd_footer.
      ENDIF.

    ENDIF.
* ----------------------------------------------------------------------
    IF ep_even_header IS SUPPLIED.

      IF me->even_header-left_value IS NOT INITIAL.
        lv_value = me->process_header_footer( ip_header = me->even_header ip_side = 'LEFT' ).
        CONCATENATE lc_marker_left lv_value INTO ep_even_header.
      ENDIF.

      IF me->even_header-center_value IS NOT INITIAL.
        lv_value = me->process_header_footer( ip_header = me->even_header ip_side = 'CENTER' ).
        CONCATENATE ep_even_header lc_marker_center lv_value INTO ep_even_header.
      ENDIF.

      IF me->even_header-right_value IS NOT INITIAL.
        lv_value = me->process_header_footer( ip_header = me->even_header ip_side = 'RIGHT' ).
        CONCATENATE ep_even_header lc_marker_right lv_value INTO ep_even_header.
      ENDIF.

      IF me->even_header-left_image IS NOT INITIAL.
        lv_value = '&G'.
        CONCATENATE ep_odd_header lc_marker_left lv_value INTO ep_even_header.
      ENDIF.
      IF me->even_header-center_image IS NOT INITIAL.
        lv_value = '&G'.
        CONCATENATE ep_odd_header lc_marker_center lv_value INTO ep_even_header.
      ENDIF.
      IF me->even_header-right_image IS NOT INITIAL.
        lv_value = '&G'.
        CONCATENATE ep_odd_header lc_marker_right lv_value INTO ep_even_header.
      ENDIF.

    ENDIF.
* ----------------------------------------------------------------------
    IF ep_even_footer IS SUPPLIED.

      IF me->even_footer-left_value IS NOT INITIAL.
        lv_value = me->process_header_footer( ip_header = me->even_footer ip_side = 'LEFT' ).
        CONCATENATE lc_marker_left lv_value INTO ep_even_footer.
      ENDIF.

      IF me->even_footer-center_value IS NOT INITIAL.
        lv_value = me->process_header_footer( ip_header = me->even_footer ip_side = 'CENTER' ).
        CONCATENATE ep_even_footer lc_marker_center lv_value INTO ep_even_footer.
      ENDIF.

      IF me->even_footer-right_value IS NOT INITIAL.
        lv_value = me->process_header_footer( ip_header = me->even_footer ip_side = 'RIGHT' ).
        CONCATENATE ep_even_footer lc_marker_right lv_value INTO ep_even_footer.
      ENDIF.

      IF me->even_footer-left_image IS NOT INITIAL.
        lv_value = '&G'.
        CONCATENATE ep_odd_header lc_marker_left lv_value INTO ep_even_footer.
      ENDIF.
      IF me->even_footer-center_image IS NOT INITIAL.
        lv_value = '&G'.
        CONCATENATE ep_odd_header lc_marker_center lv_value INTO ep_even_footer.
      ENDIF.
      IF me->even_footer-right_image IS NOT INITIAL.
        lv_value = '&G'.
        CONCATENATE ep_odd_header lc_marker_right lv_value INTO ep_even_footer.
      ENDIF.

    ENDIF.
* ----------------------------------------------------------------------
  ENDMETHOD.