  METHOD constructor.
    orientation = me->c_orientation_default.

* default margins
    margin_bottom = '0.75'.
    margin_footer = '0.3'.
    margin_header = '0.3'.
    margin_left   = '0.7'.
    margin_right  = '0.7'.
    margin_top    = '0.75'.

* clear page settings
    CLEAR: black_and_white,
           cell_comments,
           copies,
           draft,
           errors,
           first_page_number,
           fit_to_page,
           fit_to_height,
           fit_to_width,
           horizontal_dpi,
           orientation,
           page_order,
           paper_height,
           paper_size,
           paper_width,
           scale,
           use_first_page_num,
           use_printer_defaults,
           vertical_dpi.
  ENDMETHOD.