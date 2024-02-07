  METHOD create_xl_drawing_for_comments.
** Constant node name
    CONSTANTS: lc_xml_node_xml             TYPE string VALUE 'xml',
               lc_xml_node_ns_v            TYPE string VALUE 'urn:schemas-microsoft-com:vml',
               lc_xml_node_ns_o            TYPE string VALUE 'urn:schemas-microsoft-com:office:office',
               lc_xml_node_ns_x            TYPE string VALUE 'urn:schemas-microsoft-com:office:excel',
               " shapelayout
               lc_xml_node_shapelayout     TYPE string VALUE 'o:shapelayout',
               lc_xml_node_idmap           TYPE string VALUE 'o:idmap',
               " shapetype
               lc_xml_node_shapetype       TYPE string VALUE 'v:shapetype',
               lc_xml_node_stroke          TYPE string VALUE 'v:stroke',
               lc_xml_node_path            TYPE string VALUE 'v:path',
               " shape
               lc_xml_node_shape           TYPE string VALUE 'v:shape',
               lc_xml_node_fill            TYPE string VALUE 'v:fill',
               lc_xml_node_shadow          TYPE string VALUE 'v:shadow',
               lc_xml_node_textbox         TYPE string VALUE 'v:textbox',
               lc_xml_node_div             TYPE string VALUE 'div',
               lc_xml_node_clientdata      TYPE string VALUE 'x:ClientData',
               lc_xml_node_movewithcells   TYPE string VALUE 'x:MoveWithCells',
               lc_xml_node_sizewithcells   TYPE string VALUE 'x:SizeWithCells',
               lc_xml_node_anchor          TYPE string VALUE 'x:Anchor',
               lc_xml_node_autofill        TYPE string VALUE 'x:AutoFill',
               lc_xml_node_row             TYPE string VALUE 'x:Row',
               lc_xml_node_column          TYPE string VALUE 'x:Column',
               " attributes,
               lc_xml_attr_vext            TYPE string VALUE 'v:ext',
               lc_xml_attr_data            TYPE string VALUE 'data',
               lc_xml_attr_id              TYPE string VALUE 'id',
               lc_xml_attr_coordsize       TYPE string VALUE 'coordsize',
               lc_xml_attr_ospt            TYPE string VALUE 'o:spt',
               lc_xml_attr_joinstyle       TYPE string VALUE 'joinstyle',
               lc_xml_attr_path            TYPE string VALUE 'path',
               lc_xml_attr_gradientshapeok TYPE string VALUE 'gradientshapeok',
               lc_xml_attr_oconnecttype    TYPE string VALUE 'o:connecttype',
               lc_xml_attr_type            TYPE string VALUE 'type',
               lc_xml_attr_style           TYPE string VALUE 'style',
               lc_xml_attr_fillcolor       TYPE string VALUE 'fillcolor',
               lc_xml_attr_oinsetmode      TYPE string VALUE 'o:insetmode',
               lc_xml_attr_color           TYPE string VALUE 'color',
               lc_xml_attr_color2          TYPE string VALUE 'color2',
               lc_xml_attr_on              TYPE string VALUE 'on',
               lc_xml_attr_obscured        TYPE string VALUE 'obscured',
               lc_xml_attr_objecttype      TYPE string VALUE 'ObjectType',
               " attributes values
               lc_xml_attr_val_edit        TYPE string VALUE 'edit',
               lc_xml_attr_val_rect        TYPE string VALUE 'rect',
               lc_xml_attr_val_t           TYPE string VALUE 't',
               lc_xml_attr_val_miter       TYPE string VALUE 'miter',
               lc_xml_attr_val_auto        TYPE string VALUE 'auto',
               lc_xml_attr_val_black       TYPE string VALUE 'black',
               lc_xml_attr_val_none        TYPE string VALUE 'none',
               lc_xml_attr_val_msodir      TYPE string VALUE 'mso-direction-alt:auto',
               lc_xml_attr_val_note        TYPE string VALUE 'Note'.


    DATA: lo_document              TYPE REF TO if_ixml_document,
          lo_element_root          TYPE REF TO if_ixml_element,
          "shapelayout
          lo_element_shapelayout   TYPE REF TO if_ixml_element,
          lo_element_idmap         TYPE REF TO if_ixml_element,
          "shapetype
          lo_element_shapetype     TYPE REF TO if_ixml_element,
          lo_element_stroke        TYPE REF TO if_ixml_element,
          lo_element_path          TYPE REF TO if_ixml_element,
          "shape
          lo_element_shape         TYPE REF TO if_ixml_element,
          lo_element_fill          TYPE REF TO if_ixml_element,
          lo_element_shadow        TYPE REF TO if_ixml_element,
          lo_element_textbox       TYPE REF TO if_ixml_element,
          lo_element_div           TYPE REF TO if_ixml_element,
          lo_element_clientdata    TYPE REF TO if_ixml_element,
          lo_element_movewithcells TYPE REF TO if_ixml_element,
          lo_element_sizewithcells TYPE REF TO if_ixml_element,
          lo_element_anchor        TYPE REF TO if_ixml_element,
          lo_element_autofill      TYPE REF TO if_ixml_element,
          lo_element_row           TYPE REF TO if_ixml_element,
          lo_element_column        TYPE REF TO if_ixml_element,
          lo_iterator              TYPE REF TO zcl_excel_collection_iterator,
          lo_comments              TYPE REF TO zcl_excel_comments,
          lo_comment               TYPE REF TO zcl_excel_comment,
          lv_row                   TYPE zexcel_cell_row,
          lv_str_column            TYPE zexcel_cell_column_alpha,
          lv_column                TYPE zexcel_cell_column,
          lv_index                 TYPE i,
          lv_attr_id_index         TYPE i,
          lv_attr_id               TYPE string,
          lv_int_value             TYPE i,
          lv_int_value_string      TYPE string.
    DATA: lv_rel_id            TYPE i.


**********************************************************************
* STEP 1: Create XML document
    lo_document = me->ixml->create_document( ).

***********************************************************************
* STEP 2: Create main node relationships
    lo_element_root = lo_document->create_simple_element( name   = lc_xml_node_xml
                                                          parent = lo_document ).
    lo_element_root->set_attribute_ns( : name  = 'xmlns:v'  value = lc_xml_node_ns_v ),
                                         name  = 'xmlns:o'  value = lc_xml_node_ns_o ),
                                         name  = 'xmlns:x'  value = lc_xml_node_ns_x ).

**********************************************************************
* STEP 3: Create o:shapeLayout
* TO-DO: management of several authors
    lo_element_shapelayout = lo_document->create_simple_element( name   = lc_xml_node_shapelayout
                                                                 parent = lo_document ).

    lo_element_shapelayout->set_attribute_ns( name  = lc_xml_attr_vext
                                              value = lc_xml_attr_val_edit ).

    lo_element_idmap = lo_document->create_simple_element( name   = lc_xml_node_idmap
                                                           parent = lo_document ).
    lo_element_idmap->set_attribute_ns( : name  = lc_xml_attr_vext  value = lc_xml_attr_val_edit ),
                                          name  = lc_xml_attr_data  value = '1' ).

    lo_element_shapelayout->append_child( new_child = lo_element_idmap ).

    lo_element_root->append_child( new_child = lo_element_shapelayout ).

**********************************************************************
* STEP 4: Create v:shapetype

    lo_element_shapetype = lo_document->create_simple_element( name   = lc_xml_node_shapetype
                                                               parent = lo_document ).

    lo_element_shapetype->set_attribute_ns( : name  = lc_xml_attr_id         value = '_x0000_t202' ),
                                              name  = lc_xml_attr_coordsize  value = '21600,21600' ),
                                              name  = lc_xml_attr_ospt       value = '202' ),
                                              name  = lc_xml_attr_path       value = 'm,l,21600r21600,l21600,xe' ).

    lo_element_stroke = lo_document->create_simple_element( name   = lc_xml_node_stroke
                                                            parent = lo_document ).
    lo_element_stroke->set_attribute_ns( name  = lc_xml_attr_joinstyle       value = lc_xml_attr_val_miter ).

    lo_element_path   = lo_document->create_simple_element( name   = lc_xml_node_path
                                                            parent = lo_document ).
    lo_element_path->set_attribute_ns( : name  = lc_xml_attr_gradientshapeok value = lc_xml_attr_val_t ),
                                         name  = lc_xml_attr_oconnecttype    value = lc_xml_attr_val_rect ).

    lo_element_shapetype->append_child( : new_child = lo_element_stroke ),
                                          new_child = lo_element_path ).

    lo_element_root->append_child( new_child = lo_element_shapetype ).

**********************************************************************
* STEP 4: Create v:shapetype

    lo_comments = io_worksheet->get_comments( ).

    lo_iterator = lo_comments->get_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lv_index = sy-index.
      lo_comment ?= lo_iterator->get_next( ).

      zcl_excel_common=>convert_columnrow2column_a_row( EXPORTING i_columnrow = lo_comment->get_ref( )
                                                        IMPORTING e_column = lv_str_column
                                                                  e_row    = lv_row ).
      lv_column = zcl_excel_common=>convert_column2int( lv_str_column ).

      lo_element_shape = lo_document->create_simple_element( name   = lc_xml_node_shape
                                                             parent = lo_document ).

      lv_attr_id_index = 1024 + lv_index.
      lv_attr_id = lv_attr_id_index.
      CONCATENATE '_x0000_s' lv_attr_id INTO lv_attr_id.
      lo_element_shape->set_attribute_ns( : name  = lc_xml_attr_id          value = lv_attr_id ),
                                            name  = lc_xml_attr_type        value = '#_x0000_t202' ),
                                            name  = lc_xml_attr_style       value = 'size:auto;width:auto;height:auto;position:absolute;margin-left:117pt;margin-top:172.5pt;z-index:1;visibility:hidden' ),
                                            name  = lc_xml_attr_fillcolor   value = '#ffffe1' ),
                                            name  = lc_xml_attr_oinsetmode  value = lc_xml_attr_val_auto ).

      " Fill
      lo_element_fill = lo_document->create_simple_element( name   = lc_xml_node_fill
                                                            parent = lo_document ).
      lo_element_fill->set_attribute_ns( name = lc_xml_attr_color2  value = '#ffffe1' ).
      lo_element_shape->append_child( new_child = lo_element_fill ).
      " Shadow
      lo_element_shadow = lo_document->create_simple_element( name   = lc_xml_node_shadow
                                                              parent = lo_document ).
      lo_element_shadow->set_attribute_ns( : name = lc_xml_attr_on        value = lc_xml_attr_val_t ),
                                             name = lc_xml_attr_color     value = lc_xml_attr_val_black ),
                                             name = lc_xml_attr_obscured  value = lc_xml_attr_val_t ).
      lo_element_shape->append_child( new_child = lo_element_shadow ).
      " Path
      lo_element_path = lo_document->create_simple_element( name   = lc_xml_node_path
                                                            parent = lo_document ).
      lo_element_path->set_attribute_ns( name = lc_xml_attr_oconnecttype  value = lc_xml_attr_val_none ).
      lo_element_shape->append_child( new_child = lo_element_path ).
      " Textbox
      lo_element_textbox = lo_document->create_simple_element( name   = lc_xml_node_textbox
                                                               parent = lo_document ).
      lo_element_textbox->set_attribute_ns( name = lc_xml_attr_style  value = lc_xml_attr_val_msodir ).
      lo_element_div = lo_document->create_simple_element( name   = lc_xml_node_div
                                                           parent = lo_document ).
      lo_element_div->set_attribute_ns( name = lc_xml_attr_style  value = 'text-align:left' ).
      lo_element_textbox->append_child( new_child = lo_element_div ).
      lo_element_shape->append_child( new_child = lo_element_textbox ).
      " ClientData
      lo_element_clientdata = lo_document->create_simple_element( name   = lc_xml_node_clientdata
                                                                  parent = lo_document ).
      lo_element_clientdata->set_attribute_ns( name = lc_xml_attr_objecttype  value = lc_xml_attr_val_note ).
      lo_element_movewithcells = lo_document->create_simple_element( name   = lc_xml_node_movewithcells
                                                                     parent = lo_document ).
      lo_element_clientdata->append_child( new_child = lo_element_movewithcells ).
      lo_element_sizewithcells = lo_document->create_simple_element( name   = lc_xml_node_sizewithcells
                                                                     parent = lo_document ).
      lo_element_clientdata->append_child( new_child = lo_element_sizewithcells ).
      lo_element_anchor = lo_document->create_simple_element( name   = lc_xml_node_anchor
                                                              parent = lo_document ).
      lo_element_anchor->set_value( '2, 15, 11, 10, 4, 31, 15, 9' ).
      lo_element_clientdata->append_child( new_child = lo_element_anchor ).
      lo_element_autofill = lo_document->create_simple_element( name   = lc_xml_node_autofill
                                                                parent = lo_document ).
      lo_element_autofill->set_value( 'False' ).
      lo_element_clientdata->append_child( new_child = lo_element_autofill ).
      lo_element_row = lo_document->create_simple_element( name   = lc_xml_node_row
                                                           parent = lo_document ).
      lv_int_value = lv_row - 1.
      lv_int_value_string = lv_int_value.
      lo_element_row->set_value( lv_int_value_string ).
      lo_element_clientdata->append_child( new_child = lo_element_row ).
      lo_element_column = lo_document->create_simple_element( name   = lc_xml_node_column
                                                                parent = lo_document ).
      lv_int_value = lv_column - 1.
      lv_int_value_string = lv_int_value.
      lo_element_column->set_value( lv_int_value_string ).
      lo_element_clientdata->append_child( new_child = lo_element_column ).

      lo_element_shape->append_child( new_child = lo_element_clientdata ).

      lo_element_root->append_child( new_child = lo_element_shape ).
    ENDWHILE.

**********************************************************************
* STEP 6: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.