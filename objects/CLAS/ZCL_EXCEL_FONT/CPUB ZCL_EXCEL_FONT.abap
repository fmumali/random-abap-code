CLASS zcl_excel_font DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      BEGIN OF mty_s_font_metric,
        char       TYPE c LENGTH 1,
        char_width TYPE tdcwidths,
      END OF mty_s_font_metric .
    TYPES:
      mty_th_font_metrics
               TYPE HASHED TABLE OF mty_s_font_metric
               WITH UNIQUE KEY char .
    TYPES:
      BEGIN OF mty_s_font_cache,
        font_name       TYPE zexcel_style_font_name,
        font_height     TYPE tdfontsize,
        flag_bold       TYPE abap_bool,
        flag_italic     TYPE abap_bool,
        th_font_metrics TYPE mty_th_font_metrics,
      END OF mty_s_font_cache .
    TYPES:
      mty_th_font_cache
               TYPE HASHED TABLE OF mty_s_font_cache
               WITH UNIQUE KEY font_name font_height flag_bold flag_italic .

    CONSTANTS lc_default_font_height TYPE tdfontsize VALUE '110' ##NO_TEXT.
    CONSTANTS lc_default_font_name TYPE zexcel_style_font_name VALUE 'Calibri' ##NO_TEXT.
    CLASS-DATA mth_font_cache TYPE mty_th_font_cache .

    CLASS-METHODS calculate_text_width
      IMPORTING
        !iv_font_name   TYPE zexcel_style_font_name
        !iv_font_height TYPE tdfontsize
        !iv_flag_bold   TYPE abap_bool
        !iv_flag_italic TYPE abap_bool
        !iv_cell_value  TYPE zexcel_cell_value
      RETURNING
        VALUE(rv_width) TYPE f .