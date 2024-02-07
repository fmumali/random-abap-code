  METHOD create_content_types.


** Constant node name
    DATA: lc_xml_node_types        TYPE string VALUE 'Types',
          lc_xml_node_override     TYPE string VALUE 'Override',
          lc_xml_node_default      TYPE string VALUE 'Default',
          " Node attributes
          lc_xml_attr_partname     TYPE string VALUE 'PartName',
          lc_xml_attr_extension    TYPE string VALUE 'Extension',
          lc_xml_attr_contenttype  TYPE string VALUE 'ContentType',
          " Node namespace
          lc_xml_node_types_ns     TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/content-types',
          " Node extension
          lc_xml_node_rels_ext     TYPE string VALUE 'rels',
          lc_xml_node_xml_ext      TYPE string VALUE 'xml',
          lc_xml_node_xml_vml      TYPE string VALUE 'vml',   " (+) GGAR
          " Node partnumber
          lc_xml_node_theme_pn     TYPE string VALUE '/xl/theme/theme1.xml',
          lc_xml_node_styles_pn    TYPE string VALUE '/xl/styles.xml',
          lc_xml_node_workb_pn     TYPE string VALUE '/xl/workbook.xml',
          lc_xml_node_props_pn     TYPE string VALUE '/docProps/app.xml',
          lc_xml_node_worksheet_pn TYPE string VALUE '/xl/worksheets/sheet#.xml',
          lc_xml_node_strings_pn   TYPE string VALUE '/xl/sharedStrings.xml',
          lc_xml_node_core_pn      TYPE string VALUE '/docProps/core.xml',
          lc_xml_node_chart_pn     TYPE string VALUE '/xl/charts/chart#.xml',
          " Node contentType
          lc_xml_node_theme_ct     TYPE string VALUE 'application/vnd.openxmlformats-officedocument.theme+xml',
          lc_xml_node_styles_ct    TYPE string VALUE 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
          lc_xml_node_workb_ct     TYPE string VALUE 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
          lc_xml_node_rels_ct      TYPE string VALUE 'application/vnd.openxmlformats-package.relationships+xml',
          lc_xml_node_vml_ct       TYPE string VALUE 'application/vnd.openxmlformats-officedocument.vmlDrawing',
          lc_xml_node_xml_ct       TYPE string VALUE 'application/xml',
          lc_xml_node_props_ct     TYPE string VALUE 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
          lc_xml_node_worksheet_ct TYPE string VALUE 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
          lc_xml_node_strings_ct   TYPE string VALUE 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
          lc_xml_node_core_ct      TYPE string VALUE 'application/vnd.openxmlformats-package.core-properties+xml',
          lc_xml_node_table_ct     TYPE string VALUE 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
          lc_xml_node_comments_ct  TYPE string VALUE 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',   " (+) GGAR
          lc_xml_node_drawings_ct  TYPE string VALUE 'application/vnd.openxmlformats-officedocument.drawing+xml',
          lc_xml_node_chart_ct     TYPE string VALUE 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'.

    DATA: lo_document        TYPE REF TO if_ixml_document,
          lo_element_root    TYPE REF TO if_ixml_element,
          lo_element         TYPE REF TO if_ixml_element,
          lo_worksheet       TYPE REF TO zcl_excel_worksheet,
          lo_iterator        TYPE REF TO zcl_excel_collection_iterator,
          lo_nested_iterator TYPE REF TO zcl_excel_collection_iterator,
          lo_table           TYPE REF TO zcl_excel_table.

    DATA: lv_worksheets_num        TYPE i,
          lv_worksheets_numc       TYPE n LENGTH 3,
          lv_xml_node_worksheet_pn TYPE string,
          lv_value                 TYPE string,
          lv_comment_index         TYPE i VALUE 1,  " (+) GGAR
          lv_drawing_index         TYPE i VALUE 1,
          lv_index_str             TYPE string.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node types
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_types
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
    value = lc_xml_node_types_ns ).

**********************************************************************
* STEP 4: Create subnodes

    " rels node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                  value = lc_xml_node_rels_ext ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_rels_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " extension node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                  value = lc_xml_node_xml_ext ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_xml_ct ).
    lo_element_root->append_child( new_child = lo_element ).

* Begin - Add - GGAR
    " VML node (for comments)
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                  value = lc_xml_node_xml_vml ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_vml_ct ).
    lo_element_root->append_child( new_child = lo_element ).
* End   - Add - GGAR

    " Theme node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_theme_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_theme_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Styles node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_styles_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_styles_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Workbook node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_workb_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_workb_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Properties node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_props_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_props_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Worksheet node
    lv_worksheets_num = excel->get_worksheets_size( ).
    DO lv_worksheets_num TIMES.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                       parent = lo_document ).

      lv_worksheets_numc = sy-index.
      SHIFT lv_worksheets_numc LEFT DELETING LEADING '0'.
      lv_xml_node_worksheet_pn = lc_xml_node_worksheet_pn.
      REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_worksheet_pn WITH lv_worksheets_numc.
      lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                    value = lv_xml_node_worksheet_pn ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                    value = lc_xml_node_worksheet_ct ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDDO.

    lo_iterator = me->excel->get_worksheets_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).

      lo_nested_iterator = lo_worksheet->get_tables_iterator( ).

      WHILE lo_nested_iterator->has_next( ) EQ abap_true.
        lo_table ?= lo_nested_iterator->get_next( ).

        lv_value = lo_table->get_name( ).
        CONCATENATE '/xl/tables/' lv_value '.xml' INTO lv_value.

        lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lv_value ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_table_ct ).
        lo_element_root->append_child( new_child = lo_element ).
      ENDWHILE.

* Begin - Add - GGAR
      " Comments
      DATA: lo_comments TYPE REF TO zcl_excel_comments.

      lo_comments = lo_worksheet->get_comments( ).
      IF lo_comments->is_empty( ) = abap_false.
        lv_index_str = lv_comment_index.
        CONDENSE lv_index_str NO-GAPS.
        CONCATENATE '/' me->c_xl_comments INTO lv_value.
        REPLACE '#' WITH lv_index_str INTO lv_value.

        lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                         parent = lo_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                      value = lv_value ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                      value = lc_xml_node_comments_ct ).
        lo_element_root->append_child( new_child = lo_element ).

        ADD 1 TO lv_comment_index.
      ENDIF.
* End   - Add - GGAR

      " Drawings
      DATA: lo_drawings TYPE REF TO zcl_excel_drawings.

      lo_drawings = lo_worksheet->get_drawings( ).
      IF lo_drawings->is_empty( ) = abap_false.
        lv_index_str = lv_drawing_index.
        CONDENSE lv_index_str NO-GAPS.
        CONCATENATE '/' me->c_xl_drawings INTO lv_value.
        REPLACE '#' WITH lv_index_str INTO lv_value.

        lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lv_value ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_drawings_ct ).
        lo_element_root->append_child( new_child = lo_element ).

        ADD 1 TO lv_drawing_index.
      ENDIF.
    ENDWHILE.

    " media mimes
    DATA: lo_drawing    TYPE REF TO zcl_excel_drawing,
          lt_media_type TYPE TABLE OF mimetypes-extension,
          lv_media_type TYPE mimetypes-extension,
          lv_mime_type  TYPE mimetypes-type.

    lo_iterator = me->excel->get_drawings_iterator( zcl_excel_drawing=>type_image ).
    WHILE lo_iterator->has_next( ) = abap_true.
      lo_drawing ?= lo_iterator->get_next( ).

      lv_media_type = lo_drawing->get_media_type( ).
      COLLECT lv_media_type INTO lt_media_type.
    ENDWHILE.

    LOOP AT lt_media_type INTO lv_media_type.
      CALL FUNCTION 'SDOK_MIMETYPE_GET'
        EXPORTING
          extension = lv_media_type
        IMPORTING
          mimetype  = lv_mime_type.

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                       parent = lo_document ).
      lv_value = lv_media_type.
      lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                    value = lv_value ).
      lv_value = lv_mime_type.
      lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                    value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDLOOP.

    " Charts
    lo_iterator = me->excel->get_drawings_iterator( zcl_excel_drawing=>type_chart ).
    WHILE lo_iterator->has_next( ) = abap_true.
      lo_drawing ?= lo_iterator->get_next( ).

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                       parent = lo_document ).
      lv_index_str = lo_drawing->get_index( ).
      CONDENSE lv_index_str.
      lv_value = lc_xml_node_chart_pn.
      REPLACE ALL OCCURRENCES OF '#' IN lv_value WITH lv_index_str.
      lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                    value = lc_xml_node_chart_ct ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDWHILE.

    " Strings node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_strings_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_strings_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Strings node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_core_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_core_ct ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  ENDMETHOD.