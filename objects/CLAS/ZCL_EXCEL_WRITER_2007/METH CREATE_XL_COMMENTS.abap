  METHOD create_xl_comments.
** Constant node name
    CONSTANTS: lc_xml_node_comments    TYPE string VALUE 'comments',
               lc_xml_node_ns          TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
               " authors
               lc_xml_node_author      TYPE string VALUE 'author',
               lc_xml_node_authors     TYPE string VALUE 'authors',
               " comments
               lc_xml_node_commentlist TYPE string VALUE 'commentList',
               lc_xml_node_comment     TYPE string VALUE 'comment',
               lc_xml_node_text        TYPE string VALUE 'text',
               lc_xml_node_r           TYPE string VALUE 'r',
               lc_xml_node_rpr         TYPE string VALUE 'rPr',
               lc_xml_node_b           TYPE string VALUE 'b',
               lc_xml_node_sz          TYPE string VALUE 'sz',
               lc_xml_node_color       TYPE string VALUE 'color',
               lc_xml_node_rfont       TYPE string VALUE 'rFont',
*             lc_xml_node_charset     TYPE string VALUE 'charset',
               lc_xml_node_family      TYPE string VALUE 'family',
               lc_xml_node_t           TYPE string VALUE 't',
               " comments attributes
               lc_xml_attr_ref         TYPE string VALUE 'ref',
               lc_xml_attr_authorid    TYPE string VALUE 'authorId',
               lc_xml_attr_val         TYPE string VALUE 'val',
               lc_xml_attr_indexed     TYPE string VALUE 'indexed',
               lc_xml_attr_xmlspacing  TYPE string VALUE 'xml:space'.


    DATA: lo_document            TYPE REF TO if_ixml_document,
          lo_element_root        TYPE REF TO if_ixml_element,
          lo_element_authors     TYPE REF TO if_ixml_element,
          lo_element_author      TYPE REF TO if_ixml_element,
          lo_element_commentlist TYPE REF TO if_ixml_element,
          lo_element_comment     TYPE REF TO if_ixml_element,
          lo_element_text        TYPE REF TO if_ixml_element,
          lo_element_r           TYPE REF TO if_ixml_element,
          lo_element_rpr         TYPE REF TO if_ixml_element,
          lo_element_b           TYPE REF TO if_ixml_element,
          lo_element_sz          TYPE REF TO if_ixml_element,
          lo_element_color       TYPE REF TO if_ixml_element,
          lo_element_rfont       TYPE REF TO if_ixml_element,
*       lo_element_charset     TYPE REF TO if_ixml_element,
          lo_element_family      TYPE REF TO if_ixml_element,
          lo_element_t           TYPE REF TO if_ixml_element,
          lo_iterator            TYPE REF TO zcl_excel_collection_iterator,
          lo_comments            TYPE REF TO zcl_excel_comments,
          lo_comment             TYPE REF TO zcl_excel_comment.
    DATA: lv_rel_id TYPE i,
          lv_author TYPE string.


**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root = lo_document->create_simple_element( name   = lc_xml_node_comments
                                                          parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_ns ).

**********************************************************************
* STEP 4: Create authors
* TO-DO: management of several authors
    lo_element_authors = lo_document->create_simple_element( name   = lc_xml_node_authors
                                                             parent = lo_document ).

    lo_element_author  = lo_document->create_simple_element( name   = lc_xml_node_author
                                                             parent = lo_document ).
    lv_author = sy-uname.
    lo_element_author->set_value( lv_author ).

    lo_element_authors->append_child( new_child = lo_element_author ).
    lo_element_root->append_child( new_child = lo_element_authors ).

**********************************************************************
* STEP 5: Create comments

    lo_element_commentlist = lo_document->create_simple_element( name   = lc_xml_node_commentlist
                                                                 parent = lo_document ).

    lo_comments = io_worksheet->get_comments( ).

    lo_iterator = lo_comments->get_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_comment ?= lo_iterator->get_next( ).

      lo_element_comment = lo_document->create_simple_element( name   = lc_xml_node_comment
                                                               parent = lo_document ).
      lo_element_comment->set_attribute_ns( name  = lc_xml_attr_ref
                                            value = lo_comment->get_ref( ) ).
      lo_element_comment->set_attribute_ns( name  = lc_xml_attr_authorid
                                            value = '0' ).  " TO-DO

      lo_element_text = lo_document->create_simple_element( name   = lc_xml_node_text
                                                            parent = lo_document ).
      lo_element_r    = lo_document->create_simple_element( name   = lc_xml_node_r
                                                            parent = lo_document ).
      lo_element_rpr  = lo_document->create_simple_element( name   = lc_xml_node_rpr
                                                            parent = lo_document ).

      lo_element_b    = lo_document->create_simple_element( name   = lc_xml_node_b
                                                            parent = lo_document ).
      lo_element_rpr->append_child( new_child = lo_element_b ).

      add_1_val_child_node( io_document = lo_document io_parent = lo_element_rpr iv_elem_name = lc_xml_node_sz     iv_attr_name = lc_xml_attr_val     iv_attr_value = '9' ).
      add_1_val_child_node( io_document = lo_document io_parent = lo_element_rpr iv_elem_name = lc_xml_node_color  iv_attr_name = lc_xml_attr_indexed iv_attr_value = '81' ).
      add_1_val_child_node( io_document = lo_document io_parent = lo_element_rpr iv_elem_name = lc_xml_node_rfont  iv_attr_name = lc_xml_attr_val     iv_attr_value = 'Tahoma' ).
      add_1_val_child_node( io_document = lo_document io_parent = lo_element_rpr iv_elem_name = lc_xml_node_family iv_attr_name = lc_xml_attr_val     iv_attr_value = '2' ).

      lo_element_r->append_child( new_child = lo_element_rpr ).

      lo_element_t    = lo_document->create_simple_element( name   = lc_xml_node_t
                                                            parent = lo_document ).
      lo_element_t->set_attribute_ns( name  = lc_xml_attr_xmlspacing
                                      value = 'preserve' ).
      lo_element_t->set_value( lo_comment->get_text( ) ).
      lo_element_r->append_child( new_child = lo_element_t ).

      lo_element_text->append_child( new_child = lo_element_r ).
      lo_element_comment->append_child( new_child = lo_element_text ).
      lo_element_commentlist->append_child( new_child = lo_element_comment ).
    ENDWHILE.

    lo_element_root->append_child( new_child = lo_element_commentlist ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.