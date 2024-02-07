  METHOD load_worksheet_drawing.

    TYPES: BEGIN OF t_c_nv_pr,
             name TYPE string,
             id   TYPE string,
           END OF t_c_nv_pr.

    TYPES: BEGIN OF t_blip,
             cstate TYPE string,
             embed  TYPE string,
           END OF t_blip.

    TYPES: BEGIN OF t_chart,
             id TYPE string,
           END OF t_chart.

    CONSTANTS: lc_xml_attr_true     TYPE string VALUE 'true',
               lc_xml_attr_true_int TYPE string VALUE '1'.
    CONSTANTS: lc_rel_chart TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
               lc_rel_image TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'.

    DATA: drawing           TYPE REF TO if_ixml_document,
          anchors           TYPE REF TO if_ixml_node_collection,
          node              TYPE REF TO if_ixml_element,
          coll_length       TYPE i,
          iterator          TYPE REF TO if_ixml_node_iterator,
          anchor_elem       TYPE REF TO if_ixml_element,

          relationship      TYPE t_relationship,
          rel_drawings      TYPE t_rel_drawings,
          rel_drawing       TYPE t_rel_drawing,
          rels_drawing      TYPE REF TO if_ixml_document,
          rels_drawing_path TYPE string,
          stripped_name     TYPE chkfile,
          dirname           TYPE string,

          path              TYPE string,
          path2             TYPE text255,
          file_ext2         TYPE char10.

    " Read Workbook Relationships
    CALL FUNCTION 'TRINT_SPLIT_FILE_AND_PATH'
      EXPORTING
        full_name     = ip_path
      IMPORTING
        stripped_name = stripped_name
        file_path     = dirname.
    CONCATENATE dirname '_rels/' stripped_name '.rels'
      INTO rels_drawing_path.
    rels_drawing_path = resolve_path( rels_drawing_path ).
    rels_drawing = me->get_ixml_from_zip_archive( rels_drawing_path ).
    node ?= rels_drawing->find_from_name_ns( name = 'Relationship' uri = namespace-relationships ).
    WHILE node IS BOUND.
      fill_struct_from_attributes( EXPORTING ip_element = node CHANGING cp_structure = relationship ).

      rel_drawing-id = relationship-id.

      CONCATENATE dirname relationship-target INTO path.
      path = resolve_path( path ).
      rel_drawing-content = me->get_from_zip_archive( path ). "------------> This is for template usage

      path2 = path.
      zcl_excel_common=>split_file( EXPORTING ip_file = path2
                                    IMPORTING ep_extension = file_ext2 ).
      rel_drawing-file_ext = file_ext2.

      "-------------Added by Alessandro Iannacci - Should load graph xml
      CASE relationship-type.
        WHEN lc_rel_chart.
          "Read chart xml
          rel_drawing-content_xml = me->get_ixml_from_zip_archive( path ).
        WHEN OTHERS.
      ENDCASE.
      "----------------------------


      APPEND rel_drawing TO rel_drawings.

      node ?= node->get_next( ).
    ENDWHILE.

    drawing = me->get_ixml_from_zip_archive( ip_path ).

* one-cell anchor **************
    anchors = drawing->get_elements_by_tag_name_ns( name = 'oneCellAnchor' uri = namespace-xdr ).
    coll_length = anchors->get_length( ).
    iterator = anchors->create_iterator( ).
    DO coll_length TIMES.
      anchor_elem ?= iterator->get_next( ).

      CALL METHOD me->load_drawing_anchor
        EXPORTING
          io_anchor_element   = anchor_elem
          io_worksheet        = io_worksheet
          it_related_drawings = rel_drawings.

    ENDDO.

* two-cell anchor ******************
    anchors = drawing->get_elements_by_tag_name_ns( name = 'twoCellAnchor' uri = namespace-xdr ).
    coll_length = anchors->get_length( ).
    iterator = anchors->create_iterator( ).
    DO coll_length TIMES.
      anchor_elem ?= iterator->get_next( ).

      CALL METHOD me->load_drawing_anchor
        EXPORTING
          io_anchor_element   = anchor_elem
          io_worksheet        = io_worksheet
          it_related_drawings = rel_drawings.

    ENDDO.

  ENDMETHOD.