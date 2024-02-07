  METHOD itab_to_xstring.
FIELD-SYMBOLS: <fS_DAT> TYPE ANY TABLE.
CLEAR RV_XSTRING.
ASSIGN ir_data_ref->* to <fS_DAT>.
try.
cl_Salv_table=>factory( IMPORTING r_salv_table  = DATA(lr_table)
                        CHANGING  t_table       = <fS_DAT> ).

data(lt_fcat) = cl_Salv_controller_metadata=>get_lvc_fieldcatalog( r_columns = lr_table->get_columns(  )
                                                                    r_aggregations = lr_table->get_aggregations(  ) ).

data(lr_result) = cl_salv_Ex_util=>factory_result_data_table( r_Data          = ir_data_ref
                                                              t_fieldcatalog  = lt_fcat ).

cl_salv_bs_tt_util=>if_salv_bs_tt_util~transform(
          EXPORTING
            xml_type      = if_salv_bs_xml=>c_type_xlsx
            xml_version   = cl_salv_bs_a_xml_base=>get_version( )
            r_result_data = lr_result
            xml_flavour   = if_salv_bs_c_tt=>c_tt_xml_flavour_export
            gui_type      = if_salv_bs_xml=>c_gui_type_gui
          IMPORTING
            xml           = rv_xstring ).
      CATCH cx_root.
        CLEAR rv_xstring.
    ENDTRY.
  ENDMETHOD.