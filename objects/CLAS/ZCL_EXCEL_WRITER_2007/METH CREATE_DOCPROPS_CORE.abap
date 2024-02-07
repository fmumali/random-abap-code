  METHOD create_docprops_core.


** Constant node name
    DATA: lc_xml_node_coreproperties TYPE string VALUE 'coreProperties',
          lc_xml_node_creator        TYPE string VALUE 'creator',
          lc_xml_node_description    TYPE string VALUE 'description',
          lc_xml_node_lastmodifiedby TYPE string VALUE 'lastModifiedBy',
          lc_xml_node_created        TYPE string VALUE 'created',
          lc_xml_node_modified       TYPE string VALUE 'modified',
          " Node attributes
          lc_xml_attr_type           TYPE string VALUE 'type',
          lc_xml_attr_target         TYPE string VALUE 'dcterms:W3CDTF',
          " Node namespace
          lc_cp_ns                   TYPE string VALUE 'cp',
          lc_dc_ns                   TYPE string VALUE 'dc',
          lc_dcterms_ns              TYPE string VALUE 'dcterms',
*        lc_dcmitype_ns              TYPE string VALUE 'dcmitype',
          lc_xsi_ns                  TYPE string VALUE 'xsi',
          lc_xml_node_cp_ns          TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
          lc_xml_node_dc_ns          TYPE string VALUE 'http://purl.org/dc/elements/1.1/',
          lc_xml_node_dcterms_ns     TYPE string VALUE 'http://purl.org/dc/terms/',
          lc_xml_node_dcmitype_ns    TYPE string VALUE 'http://purl.org/dc/dcmitype/',
          lc_xml_node_xsi_ns         TYPE string VALUE 'http://www.w3.org/2001/XMLSchema-instance'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element.

    DATA: lv_value TYPE string,
          lv_date  TYPE d,
          lv_time  TYPE t.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node coreProperties
    lo_element_root  = lo_document->create_simple_element_ns( name   = lc_xml_node_coreproperties
                                                              prefix = lc_cp_ns
                                                              parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:cp'
                                       value = lc_xml_node_cp_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:dc'
                                       value = lc_xml_node_dc_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:dcterms'
                                       value = lc_xml_node_dcterms_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:dcmitype'
                                       value = lc_xml_node_dcmitype_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:xsi'
                                       value = lc_xml_node_xsi_ns ).

**********************************************************************
* STEP 4: Create subnodes
    " Creator node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_creator
                                                        prefix = lc_dc_ns
                                                        parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~creator.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " Description node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_description
                                                        prefix = lc_dc_ns
                                                        parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~description.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " lastModifiedBy node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_lastmodifiedby
                                                        prefix = lc_cp_ns
                                                        parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~lastmodifiedby.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " Created node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_created
                                                        prefix = lc_dcterms_ns
                                                        parent = lo_document ).
    lo_element->set_attribute_ns( name    = lc_xml_attr_type
                                  prefix  = lc_xsi_ns
                                  value   = lc_xml_attr_target ).

    CONVERT TIME STAMP excel->zif_excel_book_properties~created TIME ZONE sy-zonlo INTO DATE lv_date TIME lv_time.
    CONCATENATE lv_date lv_time INTO lv_value RESPECTING BLANKS.
    REPLACE ALL OCCURRENCES OF REGEX '([0-9]{4})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})' IN lv_value WITH '$1-$2-$3T$4:$5:$6Z'.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " Modified node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_modified
                                                        prefix = lc_dcterms_ns
                                                        parent = lo_document ).
    lo_element->set_attribute_ns( name    = lc_xml_attr_type
                                  prefix  = lc_xsi_ns
                                  value   = lc_xml_attr_target ).
    CONVERT TIME STAMP excel->zif_excel_book_properties~modified TIME ZONE sy-zonlo INTO DATE lv_date TIME lv_time.
    CONCATENATE lv_date lv_time INTO lv_value RESPECTING BLANKS.
    REPLACE ALL OCCURRENCES OF REGEX '([0-9]{4})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})' IN lv_value WITH '$1-$2-$3T$4:$5:$6Z'.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  ENDMETHOD.