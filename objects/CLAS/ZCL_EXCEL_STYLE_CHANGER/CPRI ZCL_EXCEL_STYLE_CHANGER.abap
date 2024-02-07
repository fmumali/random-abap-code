  PRIVATE SECTION.

    METHODS clear_initial_colorxfields
      IMPORTING
        is_color  TYPE zexcel_s_style_color
      CHANGING
        cs_xcolor TYPE zexcel_s_cstylex_color.

    METHODS move_supplied_borders
      IMPORTING
        iv_border_supplied        TYPE abap_bool
        is_border                 TYPE zexcel_s_cstyle_border
        iv_xborder_supplied       TYPE abap_bool
        is_xborder                TYPE zexcel_s_cstylex_border
      CHANGING
        cs_complete_style_border  TYPE zexcel_s_cstyle_border
        cs_complete_stylex_border TYPE zexcel_s_cstylex_border.

    DATA: excel                   TYPE REF TO zcl_excel,
          lv_xborder_supplied     TYPE abap_bool,
          single_change_requested TYPE zexcel_s_cstylex_complete,
          BEGIN OF multiple_change_requested,
            complete   TYPE abap_bool,
            font       TYPE abap_bool,
            fill       TYPE abap_bool,
            BEGIN OF borders,
              complete   TYPE abap_bool,
              allborders TYPE abap_bool,
              diagonal   TYPE abap_bool,
              down       TYPE abap_bool,
              left       TYPE abap_bool,
              right      TYPE abap_bool,
              top        TYPE abap_bool,
            END OF borders,
            alignment  TYPE abap_bool,
            protection TYPE abap_bool,
          END OF multiple_change_requested.
    CONSTANTS:
          lv_border_supplied  TYPE abap_bool VALUE abap_true.
    ALIASES:
          complete_style   FOR zif_excel_style_changer~complete_style,
          complete_stylex  FOR zif_excel_style_changer~complete_stylex.
