*&---------------------------------------------------------------------*
*& Report zdemo_factory_optimized_alv
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zdemo_factory_optimized_alv.

SELECT * FROM sbook INTO TABLE @DATA(bookings).

TRY.
  "Create ALV table object for the output data table
  cl_salv_table=>factory( IMPORTING r_salv_table = DATA(lr_tab)
                          CHANGING  t_table      = bookings ).
  lr_tab->get_functions( )->set_all( abap_true ).
  lr_tab->get_columns( )->set_optimize( ).
  lr_tab->get_display_settings( )->set_list_header( 'BOOKINGS' ).
  lr_tab->get_display_settings( )->set_striped_pattern( abap_true ).

  lr_tab->display( ).

  CATCH cx_root.
    MESSAGE 'Error in ALV creation' TYPE 'E'.
ENDTRY.