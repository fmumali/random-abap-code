  PROTECTED SECTION.

    DATA wt_sort TYPE lvc_t_sort .
    DATA wt_filt TYPE lvc_t_filt .
    DATA wt_fcat TYPE lvc_t_fcat .
    DATA ws_layo TYPE lvc_s_layo .
    DATA ws_option TYPE zexcel_s_converter_option .

    METHODS update_catalog
      CHANGING
        !cs_layout       TYPE zexcel_s_converter_layo
        !ct_fieldcatalog TYPE zexcel_t_converter_fcat .
    METHODS apply_sort
      IMPORTING
        !it_table TYPE STANDARD TABLE
      EXPORTING
        !eo_table TYPE REF TO data .
    METHODS get_color
      IMPORTING
        !io_table  TYPE REF TO data
      EXPORTING
        !et_colors TYPE zexcel_t_converter_col .
    METHODS get_filter
      EXPORTING
        !et_filter TYPE zexcel_t_converter_fil
      CHANGING
        !xo_table  TYPE REF TO data .
*"* private components of class ZCL_EXCEL_CONVERTER_ALV
*"* do not include other source files here!!!