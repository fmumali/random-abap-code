  PRIVATE SECTION.

    DATA autofilters TYPE REF TO zcl_excel_autofilters .
    DATA charts TYPE REF TO zcl_excel_drawings .
    DATA default_style TYPE zexcel_cell_style .
*"* private components of class ZCL_EXCEL
*"* do not include other source files here!!!
    DATA drawings TYPE REF TO zcl_excel_drawings .
    DATA ranges TYPE REF TO zcl_excel_ranges .
    DATA styles TYPE REF TO zcl_excel_styles .
    DATA t_stylemapping1 TYPE zexcel_t_stylemapping1 .
    DATA t_stylemapping2 TYPE zexcel_t_stylemapping2 .
    DATA theme TYPE REF TO zcl_excel_theme .
    DATA comments TYPE REF TO zcl_excel_comments .

    METHODS get_style_from_guid
      IMPORTING
        !ip_guid        TYPE zexcel_cell_style
      RETURNING
        VALUE(eo_style) TYPE REF TO zcl_excel_style .
    METHODS stylemapping_dynamic_style
      IMPORTING
        !ip_style        TYPE REF TO zcl_excel_style
      RETURNING
        VALUE(eo_style2) TYPE zexcel_s_stylemapping .