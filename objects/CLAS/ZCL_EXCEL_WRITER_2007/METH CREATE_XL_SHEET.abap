  METHOD create_xl_sheet.

** Constant node name
    DATA: lc_xml_node_worksheet          TYPE string VALUE 'worksheet',
          " Node namespace
          lc_xml_node_ns                 TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          lc_xml_node_r_ns               TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          lc_xml_node_comp_ns            TYPE string VALUE 'http://schemas.openxmlformats.org/markup-compatibility/2006',
          lc_xml_node_comp_pref          TYPE string VALUE 'x14ac',
          lc_xml_node_ig_ns              TYPE string VALUE 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'.

    DATA: lo_document        TYPE REF TO if_ixml_document,
          lo_element_root    TYPE REF TO if_ixml_element,
          lo_create_xl_sheet TYPE REF TO lcl_create_xl_sheet.



**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_worksheet
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_r_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:mc'
                                       value = lc_xml_node_comp_ns ).
    lo_element_root->set_attribute_ns( name  = 'mc:Ignorable'
                                       value = lc_xml_node_comp_pref ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:x14ac'
                                       value = lc_xml_node_ig_ns ).


**********************************************************************
* STEP 4: Create subnodes

    CREATE OBJECT lo_create_xl_sheet.
    lo_create_xl_sheet->create( io_worksheet         = io_worksheet
                                io_document          = lo_document
                                iv_active            = iv_active
                                io_excel_writer_2007 = me ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.