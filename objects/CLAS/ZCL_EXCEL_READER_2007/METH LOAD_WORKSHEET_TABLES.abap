  METHOD load_worksheet_tables.

    DATA lo_ixml_table_columns TYPE REF TO if_ixml_node_collection.
    DATA lo_ixml_table_column  TYPE REF TO if_ixml_element.
    DATA lo_ixml_table TYPE REF TO if_ixml_element.
    DATA lo_ixml_table_style TYPE REF TO if_ixml_element.
    DATA lt_field_catalog TYPE zexcel_t_fieldcatalog.
    DATA ls_field_catalog TYPE zexcel_s_fieldcatalog.
    DATA lo_ixml_iterator TYPE REF TO if_ixml_node_iterator.
    DATA ls_table_settings TYPE zexcel_s_table_settings.
    DATA lv_path TYPE string.
    DATA lt_components TYPE abap_component_tab.
    DATA ls_component TYPE abap_componentdescr.
    DATA lo_rtti_table TYPE REF TO cl_abap_tabledescr.
    DATA lv_dref_table TYPE REF TO data.
    DATA lv_num_lines TYPE i.
    DATA lo_line_type TYPE REF TO cl_abap_structdescr.

    DATA: BEGIN OF ls_table,
            id             TYPE string,
            name           TYPE string,
            displayname    TYPE string,
            ref            TYPE string,
            totalsrowshown TYPE string,
          END OF ls_table.

    DATA: BEGIN OF ls_table_style,
            name              TYPE string,
            showrowstripes    TYPE string,
            showcolumnstripes TYPE string,
          END OF ls_table_style.

    DATA: BEGIN OF ls_table_column,
            id   TYPE string,
            name TYPE string,
          END OF ls_table_column.

    FIELD-SYMBOLS <ls_table> LIKE LINE OF it_tables.
    FIELD-SYMBOLS <lt_table> TYPE STANDARD TABLE.
    FIELD-SYMBOLS <ls_field> TYPE zexcel_s_fieldcatalog.

    LOOP AT it_tables ASSIGNING <ls_table>.

      CONCATENATE iv_dirname <ls_table>-target INTO lv_path.
      lv_path = resolve_path( lv_path ).

      lo_ixml_table = me->get_ixml_from_zip_archive( lv_path )->get_root_element( ).
      fill_struct_from_attributes( EXPORTING
                                     ip_element = lo_ixml_table
                                   CHANGING
                                     cp_structure = ls_table ).

      lo_ixml_table_style ?= lo_ixml_table->find_from_name( 'tableStyleInfo' ).
      fill_struct_from_attributes( EXPORTING
                                     ip_element = lo_ixml_table_style
                                   CHANGING
                                     cp_structure = ls_table_style ).

      ls_table_settings-table_name = ls_table-name.
      ls_table_settings-table_style = ls_table_style-name.
      ls_table_settings-show_column_stripes = boolc( ls_table_style-showcolumnstripes = '1' ).
      ls_table_settings-show_row_stripes = boolc( ls_table_style-showrowstripes = '1' ).

      zcl_excel_common=>convert_range2column_a_row(
        EXPORTING
          i_range        = ls_table-ref
        IMPORTING
          e_column_start = ls_table_settings-top_left_column
          e_column_end   = ls_table_settings-bottom_right_column
          e_row_start    = ls_table_settings-top_left_row
          e_row_end      = ls_table_settings-bottom_right_row ).

      lo_ixml_table_columns =  lo_ixml_table->get_elements_by_tag_name( name = 'tableColumn' ).
      lo_ixml_iterator     =  lo_ixml_table_columns->create_iterator( ).
      lo_ixml_table_column  ?= lo_ixml_iterator->get_next( ).
      CLEAR lt_field_catalog.
      WHILE lo_ixml_table_column IS BOUND.

        CLEAR ls_table_column.
        fill_struct_from_attributes( EXPORTING
                                       ip_element = lo_ixml_table_column
                                     CHANGING
                                       cp_structure = ls_table_column ).

        ls_field_catalog-position = lines( lt_field_catalog ) + 1.
        ls_field_catalog-fieldname = |COMP_{ ls_field_catalog-position PAD = '0' ALIGN = RIGHT WIDTH = 4 }|.
        ls_field_catalog-scrtext_l = ls_table_column-name.
        ls_field_catalog-dynpfld = abap_true.
        ls_field_catalog-abap_type = cl_abap_typedescr=>typekind_string.
        APPEND ls_field_catalog TO lt_field_catalog.

        lo_ixml_table_column ?= lo_ixml_iterator->get_next( ).

      ENDWHILE.

      CLEAR lt_components.
      LOOP AT lt_field_catalog ASSIGNING <ls_field>.
        CLEAR ls_component.
        ls_component-name = <ls_field>-fieldname.
        ls_component-type = cl_abap_elemdescr=>get_string( ).
        APPEND ls_component TO lt_components.
      ENDLOOP.

      lo_line_type = cl_abap_structdescr=>get( lt_components ).
      lo_rtti_table = cl_abap_tabledescr=>get( lo_line_type ).
      CREATE DATA lv_dref_table TYPE HANDLE lo_rtti_table.
      ASSIGN lv_dref_table->* TO <lt_table>.

      lv_num_lines = ls_table_settings-bottom_right_row - ls_table_settings-top_left_row.
      DO lv_num_lines TIMES.
        APPEND INITIAL LINE TO <lt_table>.
      ENDDO.

      io_worksheet->bind_table(
        EXPORTING
          ip_table            = <lt_table>
          it_field_catalog    = lt_field_catalog
          is_table_settings   = ls_table_settings ).

    ENDLOOP.

  ENDMETHOD.