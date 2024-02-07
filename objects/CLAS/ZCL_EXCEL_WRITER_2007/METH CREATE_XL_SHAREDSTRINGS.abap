  METHOD create_xl_sharedstrings.


** Constant node name
    DATA: lc_xml_node_sst         TYPE string VALUE 'sst',
          lc_xml_node_si          TYPE string VALUE 'si',
          lc_xml_node_t           TYPE string VALUE 't',
          lc_xml_node_r           TYPE string VALUE 'r',
          lc_xml_node_rpr         TYPE string VALUE 'rPr',
          " Node attributes
          lc_xml_attr_count       TYPE string VALUE 'count',
          lc_xml_attr_uniquecount TYPE string VALUE 'uniqueCount',
          " Node namespace
          lc_xml_node_ns          TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element,
          lo_sub_element  TYPE REF TO if_ixml_element,
          lo_sub2_element TYPE REF TO if_ixml_element,
          lo_font_element TYPE REF TO if_ixml_element,
          lo_iterator     TYPE REF TO zcl_excel_collection_iterator,
          lo_worksheet    TYPE REF TO zcl_excel_worksheet.

    DATA: lt_cell_data       TYPE zexcel_t_cell_data_unsorted,
          lt_cell_data_rtf   TYPE zexcel_t_cell_data_unsorted,
          lv_value           TYPE string,
          ls_shared_string   TYPE zexcel_s_shared_string,
          lv_count_str       TYPE string,
          lv_uniquecount_str TYPE string,
          lv_sytabix         TYPE i,
          lv_count           TYPE i,
          lv_uniquecount     TYPE i.

    FIELD-SYMBOLS: <fs_sheet_content> TYPE zexcel_s_cell_data,
                   <fs_rtf>           TYPE zexcel_s_rtf,
                   <fs_sheet_string>  TYPE zexcel_s_shared_string.

**********************************************************************
* STEP 1: Collect strings from each worksheet
    lo_iterator = excel->get_worksheets_iterator( ).

    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).
      APPEND LINES OF lo_worksheet->sheet_content TO lt_cell_data.
    ENDWHILE.

    DELETE lt_cell_data WHERE cell_formula IS NOT INITIAL. " delete formula content

    DESCRIBE TABLE lt_cell_data LINES lv_count.
    lv_count_str = lv_count.

    " separating plain and rich text format strings
    lt_cell_data_rtf = lt_cell_data.
    DELETE lt_cell_data WHERE rtf_tab IS NOT INITIAL.
    DELETE lt_cell_data_rtf WHERE rtf_tab IS INITIAL.

    SHIFT lv_count_str RIGHT DELETING TRAILING space.
    SHIFT lv_count_str LEFT DELETING LEADING space.

    SORT lt_cell_data BY cell_value data_type.
    DELETE ADJACENT DUPLICATES FROM lt_cell_data COMPARING cell_value data_type.

    " leave unique rich text format strings
    SORT lt_cell_data_rtf BY cell_value rtf_tab.
    DELETE ADJACENT DUPLICATES FROM lt_cell_data_rtf COMPARING cell_value rtf_tab.
    " merge into single list
    APPEND LINES OF lt_cell_data_rtf TO lt_cell_data.
    SORT lt_cell_data BY cell_value rtf_tab.
    FREE lt_cell_data_rtf.

    DESCRIBE TABLE lt_cell_data LINES lv_uniquecount.
    lv_uniquecount_str = lv_uniquecount.

    SHIFT lv_uniquecount_str RIGHT DELETING TRAILING space.
    SHIFT lv_uniquecount_str LEFT DELETING LEADING space.

    CLEAR lv_count.
    LOOP AT lt_cell_data ASSIGNING <fs_sheet_content> WHERE data_type = 's'.
      lv_sytabix = lv_count.
      ls_shared_string-string_no = lv_sytabix.
      ls_shared_string-string_value = <fs_sheet_content>-cell_value.
      ls_shared_string-string_type = <fs_sheet_content>-data_type.
      ls_shared_string-rtf_tab = <fs_sheet_content>-rtf_tab.
      INSERT ls_shared_string INTO TABLE shared_strings.
      ADD 1 TO lv_count.
    ENDLOOP.


**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_sst
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_ns ).
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_count
                                       value = lv_count_str ).
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_uniquecount
                                       value = lv_uniquecount_str ).

**********************************************************************
* STEP 4: Create subnode
    LOOP AT shared_strings ASSIGNING <fs_sheet_string>.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_si
                                                       parent = lo_document ).
      IF <fs_sheet_string>-rtf_tab IS INITIAL.
        lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_t
                                                             parent = lo_document ).
        IF boolc( contains( val = <fs_sheet_string>-string_value start = ` ` ) ) = abap_true
              OR boolc( contains( val = <fs_sheet_string>-string_value end = ` ` ) ) = abap_true.
          lo_sub_element->set_attribute( name = 'space' namespace = 'xml' value = 'preserve' ).
        ENDIF.
        lv_value = escape_string_value( <fs_sheet_string>-string_value ).
        lo_sub_element->set_value( value = lv_value ).
      ELSE.
        LOOP AT <fs_sheet_string>-rtf_tab ASSIGNING <fs_rtf>.
          lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_r
                                                               parent = lo_element ).
          TRY.
              lv_value = substring( val = <fs_sheet_string>-string_value
                                    off = <fs_rtf>-offset
                                    len = <fs_rtf>-length ).
            CATCH cx_sy_range_out_of_bounds.
              EXIT.
          ENDTRY.
          lv_value = escape_string_value( lv_value ).
          IF <fs_rtf>-font IS NOT INITIAL.
            lo_font_element = lo_document->create_simple_element( name   = lc_xml_node_rpr
                                                                  parent = lo_sub_element ).
            create_xl_styles_font_node( io_document = lo_document
                                        io_parent   = lo_font_element
                                        is_font     = <fs_rtf>-font
                                        iv_use_rtf  = abap_true ).
          ENDIF.
          lo_sub2_element = lo_document->create_simple_element( name   = lc_xml_node_t
                                                              parent = lo_sub_element ).
          IF boolc( contains( val = lv_value start = ` ` ) ) = abap_true
                OR boolc( contains( val = lv_value end = ` ` ) ) = abap_true.
            lo_sub2_element->set_attribute( name = 'space' namespace = 'xml' value = 'preserve' ).
          ENDIF.
          lo_sub2_element->set_value( lv_value ).
        ENDLOOP.
      ENDIF.
      lo_element->append_child( new_child = lo_sub_element ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDLOOP.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.