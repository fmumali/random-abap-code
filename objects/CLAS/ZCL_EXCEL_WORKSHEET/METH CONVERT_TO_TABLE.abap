  METHOD convert_to_table.

    TYPES:
      BEGIN OF ts_field_conv,
        fieldname TYPE x031l-fieldname,
        convexit  TYPE x031l-convexit,
      END OF ts_field_conv,
      BEGIN OF ts_style_conv,
        cell_style TYPE zexcel_s_cell_data-cell_style,
        abap_type  TYPE abap_typekind,
      END OF ts_style_conv.

    DATA:
      lv_row_int          TYPE zexcel_cell_row,
      lv_column_int       TYPE zexcel_cell_column,
      lv_column_alpha     TYPE zexcel_cell_column_alpha,
      lt_field_catalog    TYPE zexcel_t_fieldcatalog,
      ls_field_catalog    TYPE zexcel_s_fieldcatalog,
      lv_value            TYPE string,
      lv_maxcol           TYPE i,
      lv_maxrow           TYPE i,
      lt_field_conv       TYPE TABLE OF ts_field_conv,
      lt_comp             TYPE abap_component_tab,
      ls_comp             TYPE abap_componentdescr,
      lo_line_type        TYPE REF TO cl_abap_structdescr,
      lo_tab_type         TYPE REF TO cl_abap_tabledescr,
      lr_data             TYPE REF TO data,
      lt_comp_view        TYPE abap_component_view_tab,
      ls_comp_view        TYPE abap_simple_componentdescr,
      lt_ddic_object      TYPE dd_x031l_table,
      lt_ddic_object_comp TYPE dd_x031l_table,
      ls_ddic_object      TYPE x031l,
      lt_style_conv       TYPE TABLE OF ts_style_conv,
      ls_style_conv       TYPE ts_style_conv,
      ls_stylemapping     TYPE zexcel_s_stylemapping,
      lv_format_code      TYPE zexcel_number_format,
      lv_float            TYPE f,
      lt_map_excel_row    TYPE TABLE OF i,
      lv_index            TYPE i,
      lv_index_col        TYPE i.

    FIELD-SYMBOLS:
      <lt_data>          TYPE STANDARD TABLE,
      <ls_data>          TYPE data,
      <lv_data>          TYPE data,
      <lt_data2>         TYPE STANDARD TABLE,
      <ls_data2>         TYPE data,
      <lv_data2>         TYPE data,
      <ls_field_conv>    TYPE ts_field_conv,
      <ls_ddic_object>   TYPE x031l,
      <ls_sheet_content> TYPE zexcel_s_cell_data,
      <fs_typekind_int8> TYPE abap_typekind.

    CLEAR: et_data, er_data.

    lv_maxcol = get_highest_column( ).
    lv_maxrow = get_highest_row( ).


    " Field catalog
    lt_field_catalog = it_field_catalog.
    IF lt_field_catalog IS INITIAL.
      IF et_data IS SUPPLIED.
        lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = et_data ).
      ELSE.
        DO lv_maxcol TIMES.
          ls_field_catalog-position = sy-index.
          ls_field_catalog-fieldname = 'COL_' && sy-index.
          ls_field_catalog-dynpfld = abap_true.
          APPEND ls_field_catalog TO lt_field_catalog.
        ENDDO.
      ENDIF.
    ENDIF.

    SORT lt_field_catalog BY position.
    DELETE lt_field_catalog WHERE dynpfld NE abap_true.
    CHECK: lt_field_catalog IS NOT INITIAL.


    " Create dynamic table string columns
    ls_comp-type = cl_abap_elemdescr=>get_string( ).
    LOOP AT lt_field_catalog INTO ls_field_catalog.
      ls_comp-name = ls_field_catalog-fieldname.
      APPEND ls_comp TO lt_comp.
    ENDLOOP.
    lo_line_type = cl_abap_structdescr=>create( lt_comp ).
    lo_tab_type = cl_abap_tabledescr=>create( lo_line_type ).
    CREATE DATA er_data TYPE HANDLE lo_tab_type.
    ASSIGN er_data->* TO <lt_data>.


    " Collect field conversion rules
    IF et_data IS SUPPLIED.
*      lt_ddic_object = get_ddic_object( et_data ).
      lo_tab_type ?= cl_abap_tabledescr=>describe_by_data( et_data ).
      lo_line_type ?= lo_tab_type->get_table_line_type( ).
      lo_line_type->get_ddic_object(
        RECEIVING
          p_object = lt_ddic_object
        EXCEPTIONS
          OTHERS   = 3
      ).
      IF lt_ddic_object IS INITIAL.
        lt_comp_view = lo_line_type->get_included_view( ).
        LOOP AT lt_comp_view INTO ls_comp_view.
          ls_comp_view-type->get_ddic_object(
            RECEIVING
              p_object = lt_ddic_object_comp
            EXCEPTIONS
              OTHERS   = 3
          ).
          IF lt_ddic_object_comp IS NOT INITIAL.
            READ TABLE lt_ddic_object_comp INTO ls_ddic_object INDEX 1.
            ls_ddic_object-fieldname = ls_comp_view-name.
            APPEND ls_ddic_object TO lt_ddic_object.
          ENDIF.
        ENDLOOP.
      ENDIF.

      SORT lt_ddic_object BY fieldname.
      LOOP AT lt_field_catalog INTO ls_field_catalog.
        APPEND INITIAL LINE TO lt_field_conv ASSIGNING <ls_field_conv>.
        MOVE-CORRESPONDING ls_field_catalog TO <ls_field_conv>.
        READ TABLE lt_ddic_object ASSIGNING <ls_ddic_object> WITH KEY fieldname = <ls_field_conv>-fieldname BINARY SEARCH.
        CHECK: sy-subrc EQ 0.

        ASSIGN ('CL_ABAP_TYPEDESCR=>TYPEKIND_INT8') TO <fs_typekind_int8>.
        IF sy-subrc <> 0.
          ASSIGN space TO <fs_typekind_int8>. "not used as typekind!
        ENDIF.

        CASE <ls_ddic_object>-exid.
          WHEN cl_abap_typedescr=>typekind_int
            OR cl_abap_typedescr=>typekind_int1
            OR <fs_typekind_int8>
            OR cl_abap_typedescr=>typekind_int2
            OR cl_abap_typedescr=>typekind_packed
            OR cl_abap_typedescr=>typekind_decfloat
            OR cl_abap_typedescr=>typekind_decfloat16
            OR cl_abap_typedescr=>typekind_decfloat34
            OR cl_abap_typedescr=>typekind_float.
            " Numbers
            <ls_field_conv>-convexit = cl_abap_typedescr=>typekind_float.
          WHEN OTHERS.
            <ls_field_conv>-convexit = <ls_ddic_object>-convexit.
        ENDCASE.
      ENDLOOP.
    ENDIF.

    " Date & Time in excel style
    LOOP AT me->sheet_content ASSIGNING <ls_sheet_content> WHERE cell_style IS NOT INITIAL AND data_type IS INITIAL.
      ls_style_conv-cell_style = <ls_sheet_content>-cell_style.
      APPEND ls_style_conv TO lt_style_conv.
    ENDLOOP.
    IF lt_style_conv IS NOT INITIAL.
      SORT lt_style_conv BY cell_style.
      DELETE ADJACENT DUPLICATES FROM lt_style_conv COMPARING cell_style.

      LOOP AT lt_style_conv INTO ls_style_conv.

        ls_stylemapping = me->excel->get_style_to_guid( ls_style_conv-cell_style ).
        lv_format_code = ls_stylemapping-complete_style-number_format-format_code.
        " https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
        IF lv_format_code CS ';'.
          lv_format_code = lv_format_code(sy-fdpos).
        ENDIF.
        CHECK: lv_format_code NA '#?'.

        " Remove color pattern
        REPLACE ALL OCCURRENCES OF REGEX '\[\L[^]]*\]' IN lv_format_code WITH ''.

        IF lv_format_code CA 'yd' OR lv_format_code EQ zcl_excel_style_number_format=>c_format_date_std.
          " DATE = yyyymmdd
          ls_style_conv-abap_type = cl_abap_typedescr=>typekind_date.
        ELSEIF lv_format_code CA 'hs'.
          " TIME = hhmmss
          ls_style_conv-abap_type = cl_abap_typedescr=>typekind_time.
        ELSE.
          DELETE lt_style_conv.
          CONTINUE.
        ENDIF.

        MODIFY lt_style_conv FROM ls_style_conv TRANSPORTING abap_type.

      ENDLOOP.
    ENDIF.


*--------------------------------------------------------------------*
* Start of convert content
*--------------------------------------------------------------------*
    READ TABLE me->sheet_content TRANSPORTING NO FIELDS WITH KEY cell_row = iv_begin_row.
    IF sy-subrc EQ 0.
      lv_index = sy-tabix.
    ENDIF.

    LOOP AT me->sheet_content ASSIGNING <ls_sheet_content> FROM lv_index.
      AT NEW cell_row.
        IF iv_end_row <> 0
        AND <ls_sheet_content>-cell_row > iv_end_row.
          EXIT.
        ENDIF.
        " New line
        APPEND INITIAL LINE TO <lt_data> ASSIGNING <ls_data>.
        lv_index = sy-tabix.
      ENDAT.

      IF <ls_sheet_content>-cell_value IS NOT INITIAL.
        ASSIGN COMPONENT <ls_sheet_content>-cell_column OF STRUCTURE <ls_data> TO <lv_data>.
        IF sy-subrc EQ 0.
          " value
          <lv_data> = <ls_sheet_content>-cell_value.

          " field conversion
          READ TABLE lt_field_conv ASSIGNING <ls_field_conv> INDEX <ls_sheet_content>-cell_column.
          IF sy-subrc EQ 0 AND <ls_field_conv>-convexit IS NOT INITIAL.
            CASE <ls_field_conv>-convexit.
              WHEN cl_abap_typedescr=>typekind_float.
                lv_float = zcl_excel_common=>excel_string_to_number( <ls_sheet_content>-cell_value ).
                <lv_data> = |{ lv_float NUMBER = RAW }|.
              WHEN 'ALPHA'.
                CALL FUNCTION 'CONVERSION_EXIT_ALPHA_OUTPUT'
                  EXPORTING
                    input  = <ls_sheet_content>-cell_value
                  IMPORTING
                    output = <lv_data>.
            ENDCASE.
          ENDIF.

          " style conversion
          IF <ls_sheet_content>-cell_style IS NOT INITIAL.
            READ TABLE lt_style_conv INTO ls_style_conv WITH KEY cell_style = <ls_sheet_content>-cell_style BINARY SEARCH.
            IF sy-subrc EQ 0.
              CASE ls_style_conv-abap_type.
                WHEN cl_abap_typedescr=>typekind_date.
                  <lv_data> = zcl_excel_common=>excel_string_to_date( <ls_sheet_content>-cell_value ).
                WHEN cl_abap_typedescr=>typekind_time.
                  <lv_data> = zcl_excel_common=>excel_string_to_time( <ls_sheet_content>-cell_value ).
              ENDCASE.
            ENDIF.
          ENDIF.

          " condense
          CONDENSE <lv_data>.
        ENDIF.
      ENDIF.

      AT END OF cell_row.
        " Delete empty line
        IF <ls_data> IS INITIAL.
          DELETE <lt_data> INDEX lv_index.
        ELSE.
          APPEND <ls_sheet_content>-cell_row TO lt_map_excel_row.
        ENDIF.
      ENDAT.
    ENDLOOP.
*--------------------------------------------------------------------*
* End of convert content
*--------------------------------------------------------------------*


    IF et_data IS SUPPLIED.
*      MOVE-CORRESPONDING <lt_data> TO et_data.
      LOOP AT <lt_data> ASSIGNING <ls_data>.
        APPEND INITIAL LINE TO et_data ASSIGNING <ls_data2>.
        MOVE-CORRESPONDING <ls_data> TO <ls_data2>.
      ENDLOOP.
    ENDIF.

    " Apply conversion exit.
    LOOP AT lt_field_conv ASSIGNING <ls_field_conv>
     WHERE convexit = 'ALPHA'.
      LOOP AT et_data ASSIGNING <ls_data>.
        ASSIGN COMPONENT <ls_field_conv>-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
        CHECK: sy-subrc EQ 0 AND <lv_data> IS NOT INITIAL.
        CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
          EXPORTING
            input  = <lv_data>
          IMPORTING
            output = <lv_data>.
      ENDLOOP.
    ENDLOOP.

  ENDMETHOD.