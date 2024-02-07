  METHOD load_dxf_styles.

    DATA: lo_styles_xml   TYPE REF TO if_ixml_document,
          lo_node_dxfs    TYPE REF TO if_ixml_element,

          lo_nodes_dxf    TYPE REF TO if_ixml_node_collection,
          lo_iterator_dxf TYPE REF TO if_ixml_node_iterator,
          lo_node_dxf     TYPE REF TO if_ixml_element,

          lv_dxf_count    TYPE i.

    FIELD-SYMBOLS: <ls_dxf_style> LIKE LINE OF mt_dxf_styles.

*--------------------------------------------------------------------*
* Look for dxfs-node
*--------------------------------------------------------------------*
    lo_styles_xml = me->get_ixml_from_zip_archive( iv_path ).
    lo_node_dxfs  = lo_styles_xml->find_from_name_ns( name = 'dxfs' uri = namespace-main ).
    CHECK lo_node_dxfs IS BOUND.


*--------------------------------------------------------------------*
* loop through all dxf-nodes and create style for each
*--------------------------------------------------------------------*
    lo_nodes_dxf ?= lo_node_dxfs->get_elements_by_tag_name_ns( name = 'dxf' uri = namespace-main ).
    lo_iterator_dxf = lo_nodes_dxf->create_iterator( ).
    lo_node_dxf ?= lo_iterator_dxf->get_next( ).
    WHILE lo_node_dxf IS BOUND.

      APPEND INITIAL LINE TO mt_dxf_styles ASSIGNING <ls_dxf_style>.
      <ls_dxf_style>-dxf = lv_dxf_count. " We start counting at 0
      ADD 1 TO lv_dxf_count.             " prepare next entry

      <ls_dxf_style>-guid = get_dxf_style_guid( io_ixml_dxf = lo_node_dxf
                                                io_excel    = io_excel ).
      lo_node_dxf ?= lo_iterator_dxf->get_next( ).

    ENDWHILE.


  ENDMETHOD.