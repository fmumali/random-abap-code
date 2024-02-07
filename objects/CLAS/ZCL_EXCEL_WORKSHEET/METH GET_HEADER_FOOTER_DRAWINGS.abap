  METHOD get_header_footer_drawings.
    DATA: ls_odd_header  TYPE zexcel_s_worksheet_head_foot,
          ls_odd_footer  TYPE zexcel_s_worksheet_head_foot,
          ls_even_header TYPE zexcel_s_worksheet_head_foot,
          ls_even_footer TYPE zexcel_s_worksheet_head_foot,
          ls_hd_ft       TYPE zexcel_s_worksheet_head_foot.

    FIELD-SYMBOLS: <fs_drawings> TYPE zexcel_s_drawings.

    me->sheet_setup->get_header_footer( IMPORTING ep_odd_header = ls_odd_header
                                                  ep_odd_footer = ls_odd_footer
                                                  ep_even_header = ls_even_header
                                                  ep_even_footer = ls_even_footer ).

**********************************************************************
*** Odd header
    ls_hd_ft = ls_odd_header.
    IF ls_hd_ft-left_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-left_image.
    ENDIF.
    IF ls_hd_ft-right_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-right_image.
    ENDIF.
    IF ls_hd_ft-center_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-center_image.
    ENDIF.

**********************************************************************
*** Odd footer
    ls_hd_ft = ls_odd_footer.
    IF ls_hd_ft-left_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-left_image.
    ENDIF.
    IF ls_hd_ft-right_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-right_image.
    ENDIF.
    IF ls_hd_ft-center_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-center_image.
    ENDIF.

**********************************************************************
*** Even header
    ls_hd_ft = ls_even_header.
    IF ls_hd_ft-left_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-left_image.
    ENDIF.
    IF ls_hd_ft-right_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-right_image.
    ENDIF.
    IF ls_hd_ft-center_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-center_image.
    ENDIF.

**********************************************************************
*** Even footer
    ls_hd_ft = ls_even_footer.
    IF ls_hd_ft-left_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-left_image.
    ENDIF.
    IF ls_hd_ft-right_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-right_image.
    ENDIF.
    IF ls_hd_ft-center_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-center_image.
    ENDIF.

  ENDMETHOD.                    "get_header_footer_drawings