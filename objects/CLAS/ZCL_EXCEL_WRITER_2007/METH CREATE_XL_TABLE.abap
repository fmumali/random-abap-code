  METHOD create_xl_table.

    DATA: lc_xml_node_table        TYPE string VALUE 'table',
          lc_xml_node_relationship TYPE string VALUE 'Relationship',
          " Node attributes
          lc_xml_attr_id           TYPE string VALUE 'id',
          lc_xml_attr_name         TYPE string VALUE 'name',
          lc_xml_attr_display_name TYPE string VALUE 'displayName',
          lc_xml_attr_ref          TYPE string VALUE 'ref',
          lc_xml_attr_totals       TYPE string VALUE 'totalsRowShown',
          " Node namespace
          lc_xml_node_table_ns     TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          " Node id
          lc_xml_node_ridx_id      TYPE string VALUE 'rId#'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element,
          lo_element2     TYPE REF TO if_ixml_element,
          lo_element3     TYPE REF TO if_ixml_element,
          lv_table_name   TYPE string,
          lv_id           TYPE i,
          lv_match        TYPE i,
          lv_ref          TYPE string,
          lv_value        TYPE string,
          lv_num_columns  TYPE i,
          ls_fieldcat     TYPE zexcel_s_fieldcatalog.


**********************************************************************
* STEP 1: Create xml
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node table
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_table
                                                           parent = lo_document ).

    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_table_ns  ).

    lv_id = io_table->get_id( ).
    lv_value = zcl_excel_common=>number_to_excel_string( ip_value = lv_id ).
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_id
                                       value = lv_value ).

    FIND ALL OCCURRENCES OF REGEX '[^_a-zA-Z0-9]' IN io_table->settings-table_name IGNORING CASE MATCH COUNT lv_match.
    IF io_table->settings-table_name IS NOT INITIAL AND lv_match EQ 0.
      " Name rules (https://support.microsoft.com/en-us/office/rename-an-excel-table-fbf49a4f-82a3-43eb-8ba2-44d21233b114)
      "   - You can't use "C", "c", "R", or "r" for the name, because they're already designated as a shortcut for selecting the column or row for the active cell when you enter them in the Name or Go To box.
      "   - Don't use cell references — Names can't be the same as a cell reference, such as Z$100 or R1C1
      IF ( strlen( io_table->settings-table_name ) = 1 AND io_table->settings-table_name CO 'CcRr' )
         OR zcl_excel_common=>shift_formula(
              iv_reference_formula = io_table->settings-table_name
              iv_shift_cols        = 0
              iv_shift_rows        = 1 ) <> io_table->settings-table_name.
        lv_table_name = io_table->get_name( ).
      ELSE.
        lv_table_name = io_table->settings-table_name.
      ENDIF.
    ELSE.
      lv_table_name = io_table->get_name( ).
    ENDIF.
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_name
                                       value = lv_table_name ).

    lo_element_root->set_attribute_ns( name  = lc_xml_attr_display_name
                                       value = lv_table_name ).

    lv_ref = io_table->get_reference( ).
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_ref
                                       value = lv_ref ).
    IF io_table->has_totals( ) = abap_true.
      lo_element_root->set_attribute_ns( name  = 'totalsRowCount'
                                             value = '1' ).
    ELSE.
      lo_element_root->set_attribute_ns( name  = lc_xml_attr_totals
                                           value = '0' ).
    ENDIF.

**********************************************************************
* STEP 4: Create subnodes

    " autoFilter
    IF io_table->settings-nofilters EQ abap_false.
      lo_element = lo_document->create_simple_element( name   = 'autoFilter'
                                                       parent = lo_document ).

      lv_ref = io_table->get_reference( ip_include_totals_row = abap_false ).
      lo_element->set_attribute_ns( name  = 'ref'
                                    value = lv_ref ).

      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.

    "columns
    lo_element = lo_document->create_simple_element( name   = 'tableColumns'
                                                     parent = lo_document ).

    LOOP AT io_table->fieldcat INTO ls_fieldcat WHERE dynpfld = abap_true.
      ADD 1 TO lv_num_columns.
    ENDLOOP.

    lv_value = lv_num_columns.
    CONDENSE lv_value.
    lo_element->set_attribute_ns( name  = 'count'
                                  value = lv_value ).

    lo_element_root->append_child( new_child = lo_element ).

    LOOP AT io_table->fieldcat INTO ls_fieldcat WHERE dynpfld = abap_true.
      lo_element2 = lo_document->create_simple_element_ns( name   = 'tableColumn'
                                                                  parent = lo_element ).

      lv_value = ls_fieldcat-position.
      SHIFT lv_value LEFT DELETING LEADING '0'.
      lo_element2->set_attribute_ns( name  = 'id'
                                    value = lv_value ).

      lv_value = ls_fieldcat-column_name.

      " The text "_x...._", with "_x" not "_X", with exactly 4 ".", each being 0-9 a-f or A-F (case insensitive), is interpreted
      " like Unicode character U+.... (e.g. "_x0041_" is rendered like "A") is for characters.
      " To not interpret it, Excel replaces the first "_" is to be replaced with "_x005f_".
      IF lv_value CS '_x'.
        REPLACE ALL OCCURRENCES OF REGEX '_(x[0-9a-fA-F]{4}_)' IN lv_value WITH '_x005f_$1' RESPECTING CASE.
      ENDIF.

      " XML chapter 2.2: Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
      " NB: although Excel supports _x0009_, it's not rendered except if you edit the text.
      " Excel considers _x000d_ as being an error (_x000a_ is sufficient and rendered).
      REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>newline IN lv_value WITH '_x000a_'.
      REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf(1) IN lv_value WITH ``.
      REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>horizontal_tab IN lv_value WITH '_x0009_'.

      lo_element2->set_attribute_ns( name  = 'name'
                                    value = lv_value ).

      IF ls_fieldcat-totals_function IS NOT INITIAL.
        lo_element2->set_attribute_ns( name  = 'totalsRowFunction'
                                          value = ls_fieldcat-totals_function ).
      ENDIF.

      IF ls_fieldcat-column_formula IS NOT INITIAL.
        lv_value = ls_fieldcat-column_formula.
        CONDENSE lv_value.
        lo_element3 = lo_document->create_simple_element_ns( name   = 'calculatedColumnFormula'
                                                             parent = lo_element2 ).
        lo_element3->set_value( lv_value ).
        lo_element2->append_child( new_child = lo_element3 ).
      ENDIF.

      lo_element->append_child( new_child = lo_element2 ).
    ENDLOOP.


    lo_element = lo_document->create_simple_element( name   = 'tableStyleInfo'
                                                          parent = lo_element_root ).

    lo_element->set_attribute_ns( name  = 'name'
                                       value = io_table->settings-table_style  ).

    lo_element->set_attribute_ns( name  = 'showFirstColumn'
                                       value = '0' ).

    lo_element->set_attribute_ns( name  = 'showLastColumn'
                                       value = '0' ).

    IF io_table->settings-show_row_stripes = abap_true.
      lv_value = '1'.
    ELSE.
      lv_value = '0'.
    ENDIF.

    lo_element->set_attribute_ns( name  = 'showRowStripes'
                                       value = lv_value ).

    IF io_table->settings-show_column_stripes = abap_true.
      lv_value = '1'.
    ELSE.
      lv_value = '0'.
    ENDIF.

    lo_element->set_attribute_ns( name  = 'showColumnStripes'
                                       value = lv_value ).

    lo_element_root->append_child( new_child = lo_element ).
**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.