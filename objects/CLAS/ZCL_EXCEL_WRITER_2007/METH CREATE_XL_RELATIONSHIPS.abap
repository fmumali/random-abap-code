  METHOD create_xl_relationships.


** Constant node name
    DATA: lc_xml_node_relationships TYPE string VALUE 'Relationships',
          lc_xml_node_relationship  TYPE string VALUE 'Relationship',
          " Node attributes
          lc_xml_attr_id            TYPE string VALUE 'Id',
          lc_xml_attr_type          TYPE string VALUE 'Type',
          lc_xml_attr_target        TYPE string VALUE 'Target',
          " Node namespace
          lc_xml_node_rels_ns       TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships',
          " Node id
          lc_xml_node_ridx_id       TYPE string VALUE 'rId#',
          " Node type
          lc_xml_node_rid_sheet_tp  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
          lc_xml_node_rid_theme_tp  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
          lc_xml_node_rid_styles_tp TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
          lc_xml_node_rid_shared_tp TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
          " Node target
          lc_xml_node_ridx_tg       TYPE string VALUE 'worksheets/sheet#.xml',
          lc_xml_node_rid_shared_tg TYPE string VALUE 'sharedStrings.xml',
          lc_xml_node_rid_styles_tg TYPE string VALUE 'styles.xml',
          lc_xml_node_rid_theme_tg  TYPE string VALUE 'theme/theme1.xml'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element.

    DATA: lv_xml_node_ridx_tg TYPE string,
          lv_xml_node_ridx_id TYPE string,
          lv_size             TYPE i,
          lv_syindex          TYPE string.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_relationships
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_rels_ns ).

**********************************************************************
* STEP 4: Create subnodes

    lv_size = excel->get_worksheets_size( ).


    " Relationship node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
    parent = lo_document ).
    lv_size = lv_size + 1.
    lv_syindex = lv_size.
    SHIFT lv_syindex RIGHT DELETING TRAILING space.
    SHIFT lv_syindex LEFT DELETING LEADING space.
    lv_xml_node_ridx_id = lc_xml_node_ridx_id.
    REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_id WITH lv_syindex.
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
    value = lv_xml_node_ridx_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
    value = lc_xml_node_rid_theme_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
    value = lc_xml_node_rid_theme_tg ).
    lo_element_root->append_child( new_child = lo_element ).


    " Relationship node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lv_size = lv_size + 1.
    lv_syindex = lv_size.
    SHIFT lv_syindex RIGHT DELETING TRAILING space.
    SHIFT lv_syindex LEFT DELETING LEADING space.
    lv_xml_node_ridx_id = lc_xml_node_ridx_id.
    REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_id WITH lv_syindex.
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lv_xml_node_ridx_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid_styles_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid_styles_tg ).
    lo_element_root->append_child( new_child = lo_element ).



    lv_size = excel->get_worksheets_size( ).

    DO lv_size TIMES.
      " Relationship node
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
      parent = lo_document ).
      lv_xml_node_ridx_id = lc_xml_node_ridx_id.
      lv_xml_node_ridx_tg = lc_xml_node_ridx_tg.
      lv_syindex = sy-index.
      SHIFT lv_syindex RIGHT DELETING TRAILING space.
      SHIFT lv_syindex LEFT DELETING LEADING space.
      REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_id WITH lv_syindex.
      REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_tg WITH lv_syindex.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
      value = lv_xml_node_ridx_id ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
      value = lc_xml_node_rid_sheet_tp ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
      value = lv_xml_node_ridx_tg ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDDO.

    " Relationship node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    ADD 3 TO lv_size.
    lv_syindex = lv_size.
    SHIFT lv_syindex RIGHT DELETING TRAILING space.
    SHIFT lv_syindex LEFT DELETING LEADING space.
    lv_xml_node_ridx_id = lc_xml_node_ridx_id.
    REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_id WITH lv_syindex.
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lv_xml_node_ridx_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid_shared_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid_shared_tg ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.