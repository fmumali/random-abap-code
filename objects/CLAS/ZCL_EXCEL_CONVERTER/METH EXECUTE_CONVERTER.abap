  METHOD execute_converter.
    DATA: lo_if               TYPE REF TO zif_excel_converter,
          ls_types            TYPE ts_alv_types,
          lo_addit            TYPE REF TO cl_abap_classdescr,
          lo_addit_superclass TYPE REF TO cl_abap_classdescr.

    IF io_object IS BOUND.
      TRY.
          lo_addit            ?= cl_abap_typedescr=>describe_by_object_ref( io_object ).
        CATCH cx_sy_move_cast_error.
          RAISE EXCEPTION TYPE zcx_excel.
      ENDTRY.
      ls_types-seoclass = lo_addit->get_relative_name( ).
      READ TABLE wt_objects INTO ls_types WITH TABLE KEY seoclass = ls_types-seoclass.
      IF sy-subrc NE 0.
        DO.
          FREE lo_addit_superclass.
          lo_addit_superclass = lo_addit->get_super_class_type( ).
          IF lo_addit_superclass IS INITIAL.
            CLEAR ls_types-clsname.
            EXIT.
          ENDIF.
          lo_addit = lo_addit_superclass.
          ls_types-seoclass = lo_addit->get_relative_name( ).
          READ TABLE wt_objects INTO ls_types WITH TABLE KEY seoclass = ls_types-seoclass.
          IF sy-subrc EQ 0.
            EXIT.
          ENDIF.
        ENDDO.
      ENDIF.
      IF ls_types-clsname IS NOT INITIAL.
        CREATE OBJECT lo_if TYPE (ls_types-clsname).
        lo_if->create_fieldcatalog(
          EXPORTING
            is_option       = ws_option
            io_object       = io_object
            it_table        = it_table
          IMPORTING
            es_layout       = ws_layout
            et_fieldcatalog = wt_fieldcatalog
            eo_table        = wo_table
            et_colors       = wt_colors
            et_filter       = wt_filter
            ).
*  data lines of highest level.
        IF ws_layout-max_subtotal_level > 0. ADD 1 TO ws_layout-max_subtotal_level. ENDIF.
      ELSE.
        RAISE EXCEPTION TYPE zcx_excel.
      ENDIF.
    ELSE.
      CLEAR wt_fieldcatalog.
      GET REFERENCE OF it_table INTO wo_table.
    ENDIF.
  ENDMETHOD.