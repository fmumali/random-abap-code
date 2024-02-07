  METHOD create_docprops_app.


** Constant node name
    DATA: lc_xml_node_properties        TYPE string VALUE 'Properties',
          lc_xml_node_application       TYPE string VALUE 'Application',
          lc_xml_node_docsecurity       TYPE string VALUE 'DocSecurity',
          lc_xml_node_scalecrop         TYPE string VALUE 'ScaleCrop',
          lc_xml_node_headingpairs      TYPE string VALUE 'HeadingPairs',
          lc_xml_node_vector            TYPE string VALUE 'vector',
          lc_xml_node_variant           TYPE string VALUE 'variant',
          lc_xml_node_lpstr             TYPE string VALUE 'lpstr',
          lc_xml_node_i4                TYPE string VALUE 'i4',
          lc_xml_node_titlesofparts     TYPE string VALUE 'TitlesOfParts',
          lc_xml_node_company           TYPE string VALUE 'Company',
          lc_xml_node_linksuptodate     TYPE string VALUE 'LinksUpToDate',
          lc_xml_node_shareddoc         TYPE string VALUE 'SharedDoc',
          lc_xml_node_hyperlinkschanged TYPE string VALUE 'HyperlinksChanged',
          lc_xml_node_appversion        TYPE string VALUE 'AppVersion',
          " Namespace prefix
          lc_vt_ns                      TYPE string VALUE 'vt',
          lc_xml_node_props_ns          TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
          lc_xml_node_props_vt_ns       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
          " Node attributes
          lc_xml_attr_size              TYPE string VALUE 'size',
          lc_xml_attr_basetype          TYPE string VALUE 'baseType'.

    DATA: lo_document            TYPE REF TO if_ixml_document,
          lo_element_root        TYPE REF TO if_ixml_element,
          lo_element             TYPE REF TO if_ixml_element,
          lo_sub_element_vector  TYPE REF TO if_ixml_element,
          lo_sub_element_variant TYPE REF TO if_ixml_element,
          lo_sub_element_lpstr   TYPE REF TO if_ixml_element,
          lo_sub_element_i4      TYPE REF TO if_ixml_element,
          lo_iterator            TYPE REF TO zcl_excel_collection_iterator,
          lo_worksheet           TYPE REF TO zcl_excel_worksheet.

    DATA: lv_value                TYPE string.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node properties
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_properties
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_props_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:vt'
                                       value = lc_xml_node_props_vt_ns ).

**********************************************************************
* STEP 4: Create subnodes
    " Application
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_application
                                                     parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~application.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " DocSecurity
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_docsecurity
                                                              parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~docsecurity.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " ScaleCrop
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_scalecrop
                                                     parent = lo_document ).
    lv_value = me->flag2bool( excel->zif_excel_book_properties~scalecrop ).
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " HeadingPairs
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_headingpairs
                                                     parent = lo_document ).


    " * vector node
    lo_sub_element_vector = lo_document->create_simple_element_ns( name   = lc_xml_node_vector
                                                                   prefix = lc_vt_ns
                                                                   parent = lo_document ).
    lo_sub_element_vector->set_attribute_ns( name    = lc_xml_attr_size
                                             value   = '2' ).
    lo_sub_element_vector->set_attribute_ns( name    = lc_xml_attr_basetype
                                             value   = lc_xml_node_variant ).

    " ** variant node
    lo_sub_element_variant = lo_document->create_simple_element_ns( name   = lc_xml_node_variant
                                                                    prefix = lc_vt_ns
                                                                    parent = lo_document ).

    " *** lpstr node
    lo_sub_element_lpstr = lo_document->create_simple_element_ns( name   = lc_xml_node_lpstr
                                                                  prefix = lc_vt_ns
                                                                  parent = lo_document ).
    lv_value = excel->get_worksheets_name( ).
    lo_sub_element_lpstr->set_value( value = lv_value ).
    lo_sub_element_variant->append_child( new_child = lo_sub_element_lpstr ). " lpstr node

    lo_sub_element_vector->append_child( new_child = lo_sub_element_variant ). " variant node

    " ** variant node
    lo_sub_element_variant = lo_document->create_simple_element_ns( name   = lc_xml_node_variant
                                                                    prefix = lc_vt_ns
                                                                    parent = lo_document ).

    " *** i4 node
    lo_sub_element_i4 = lo_document->create_simple_element_ns( name   = lc_xml_node_i4
                                                               prefix = lc_vt_ns
                                                               parent = lo_document ).
    lv_value = excel->get_worksheets_size( ).
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_sub_element_i4->set_value( value = lv_value ).
    lo_sub_element_variant->append_child( new_child = lo_sub_element_i4 ). " lpstr node

    lo_sub_element_vector->append_child( new_child = lo_sub_element_variant ). " variant node

    lo_element->append_child( new_child = lo_sub_element_vector ). " vector node

    lo_element_root->append_child( new_child = lo_element ). " HeadingPairs


    " TitlesOfParts
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_titlesofparts
                                                     parent = lo_document ).


    " * vector node
    lo_sub_element_vector = lo_document->create_simple_element_ns( name   = lc_xml_node_vector
                                                                   prefix = lc_vt_ns
                                                                   parent = lo_document ).
    lv_value = excel->get_worksheets_size( ).
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_sub_element_vector->set_attribute_ns( name    = lc_xml_attr_size
                                             value   = lv_value ).
    lo_sub_element_vector->set_attribute_ns( name    = lc_xml_attr_basetype
                                             value   = lc_xml_node_lpstr ).

    lo_iterator = excel->get_worksheets_iterator( ).

    WHILE lo_iterator->has_next( ) EQ abap_true.
      " ** lpstr node
      lo_sub_element_lpstr = lo_document->create_simple_element_ns( name   = lc_xml_node_lpstr
                                                                    prefix = lc_vt_ns
                                                                    parent = lo_document ).
      lo_worksheet ?= lo_iterator->get_next( ).
      lv_value = lo_worksheet->get_title( ).
      lo_sub_element_lpstr->set_value( value = lv_value ).
      lo_sub_element_vector->append_child( new_child = lo_sub_element_lpstr ). " lpstr node
    ENDWHILE.

    lo_element->append_child( new_child = lo_sub_element_vector ). " vector node

    lo_element_root->append_child( new_child = lo_element ). " TitlesOfParts



    " Company
    IF excel->zif_excel_book_properties~company IS NOT INITIAL.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_company
                                                       parent = lo_document ).
      lv_value = excel->zif_excel_book_properties~company.
      lo_element->set_value( value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.

    " LinksUpToDate
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_linksuptodate
                                                     parent = lo_document ).
    lv_value = me->flag2bool( excel->zif_excel_book_properties~linksuptodate ).
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " SharedDoc
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_shareddoc
                                                     parent = lo_document ).
    lv_value = me->flag2bool( excel->zif_excel_book_properties~shareddoc ).
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " HyperlinksChanged
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_hyperlinkschanged
                                                     parent = lo_document ).
    lv_value = me->flag2bool( excel->zif_excel_book_properties~hyperlinkschanged ).
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " AppVersion
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_appversion
                                                     parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~appversion.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  ENDMETHOD.