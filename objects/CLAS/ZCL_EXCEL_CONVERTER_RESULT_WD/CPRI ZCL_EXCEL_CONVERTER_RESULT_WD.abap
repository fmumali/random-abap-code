  PRIVATE SECTION.

    DATA wo_config TYPE REF TO cl_salv_wd_config_table .
    DATA wt_fields TYPE salv_wd_t_field_ref .
    DATA wt_columns TYPE salv_wd_t_column_ref .

    METHODS get_columns_info
      CHANGING
        !xs_fcat TYPE lvc_s_fcat .
    METHODS get_fields_info
      CHANGING
        !xs_fcat TYPE lvc_s_fcat .
    METHODS create_wt_sort .
    METHODS create_wt_filt .
    METHODS create_wt_fcat
      IMPORTING
        !io_table TYPE REF TO data .